// aladdin.js

const STORAGE_KEY = 'aladdin_state'
const singleton = [false]

export function createAladdin(Office) {
  if (typeof window !== 'undefined' && window.aladdinInstance) return window.aladdinInstance
  if (singleton[0]) return singleton[0]
  const instance = aladdin(Office)
  if (typeof window !== 'undefined') window.aladdinInstance = instance
  if (typeof window === 'undefined') singleton[0] = instance
  instance.loadState()
  instance.watchState()
  return instance
}

function aladdin(Office) {
  return {
    Office,
    _currentItemId: null,
    _state: {
      events: [],
      currentEmail: null,
      userName: null,
      folderName: null,
      platform: null,
      version: null
    },

    state() {
      return this._state
    },

    saveState() {
      try {
        if (typeof localStorage !== 'undefined') {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(this._state))
        }
      } catch (e) {
        console.error('Failed to save state', e)
      }
    },

    loadState() {
      try {
        if (typeof localStorage !== 'undefined') {
          const stored = localStorage.getItem(STORAGE_KEY)
          if (stored) {
            const parsed = JSON.parse(stored)
            this._state.events = parsed.events || []
            this._state.currentEmail = parsed.currentEmail || null
            this._state.userName = parsed.userName || null
            this._state.folderName = parsed.folderName || null
            this._state.platform = parsed.platform || null
            this._state.version = parsed.version || null
            if (this._state.currentEmail && this._state.currentEmail.itemId) {
              this._currentItemId = this._state.currentEmail.itemId
            }
          }
        }
      } catch (e) {
        console.error('Failed to load state', e)
      }
    },

    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === STORAGE_KEY && e.newValue) {
            try {
              const parsed = JSON.parse(e.newValue)
              this._state.events = parsed.events || []
              this._state.currentEmail = parsed.currentEmail || null
              this._state.userName = parsed.userName || this._state.userName
              this._state.folderName = parsed.folderName || this._state.folderName
              this._state.platform = parsed.platform || this._state.platform
              this._state.version = parsed.version || this._state.version
              this.renderUI()
            } catch (err) {
              console.error('Failed to parse storage event', err)
            }
          }
        })
      }
    },

    event(name, details) {
      const entry = {
        name,
        details: details || {},
        timestamp: new Date().toISOString()
      }
      this._state.events.unshift(entry)
      if (this._state.events.length > 50) {
        this._state.events = this._state.events.slice(0, 50)
      }
      this.saveState()
      this.renderUI()
    },

    async initialize(userName, folderName, platform, version) {
      this._state.userName = userName
      this._state.folderName = folderName
      this._state.platform = platform
      this._state.version = version
      this.saveState()

      this.event('initialized', { userName, folderName, platform, version })

      this._registerMailboxEvents()

      await this._addMasterCategories()

      await this._captureCurrentItem()

      this.renderUI()
    },

    _registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events
      const mailboxEvents = ['ItemChanged']
      mailboxEvents.forEach((eventName) => {
        if (this.Office.EventType[eventName]) {
          try {
            mailbox.addHandlerAsync(
              this.Office.EventType[eventName],
              (eventArgs) => this._handleMailboxEvent(eventName, eventArgs),
              (result) => {
                if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                  console.log('Registered mailbox event: ' + eventName)
                } else {
                  console.warn('Failed to register mailbox event: ' + eventName, result.error)
                }
              }
            )
          } catch (e) {
            console.warn('Error registering mailbox event: ' + eventName, e)
          }
        }
      })

      // Item-level events - only if item exists
      this._registerItemEvents()
    },

    _registerItemEvents() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const itemEvents = [
        'RecipientsChanged',
        'AttachmentsChanged',
        'RecurrenceChanged',
        'AppointmentTimeChanged',
        'EnhancedLocationsChanged'
      ]

      itemEvents.forEach((eventName) => {
        if (this.Office.EventType[eventName]) {
          try {
            item.addHandlerAsync(
              this.Office.EventType[eventName],
              (eventArgs) => this._handleItemEvent(eventName, eventArgs),
              (result) => {
                if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                  console.log('Registered item event: ' + eventName)
                } else {
                  console.warn('Failed to register item event: ' + eventName, result.error)
                }
              }
            )
          } catch (e) {
            console.warn('Error registering item event: ' + eventName, e)
          }
        }
      })
    },

    _handleMailboxEvent(eventName, eventArgs) {
      this.event(eventName, { type: 'mailbox', args: eventArgs || {} })

      if (eventName === 'ItemChanged') {
        this._onItemChanged()
      }
    },

    _handleItemEvent(eventName, eventArgs) {
      this.event(eventName, { type: 'item', args: eventArgs || {} })
    },

    async _onItemChanged() {
      // Notify API about the previously selected email before capturing new one
      if (this._state.currentEmail) {
        await this.notify()
      }

      // Re-register item-level events for the new item
      this._registerItemEvents()

      // Capture the newly selected item
      await this._captureCurrentItem()

      // Try to update folder info from the new item's headers
      await this._updateFolderFromHeaders()

      this.renderUI()
    },

    async _captureCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) {
        this._currentItemId = null
        this._state.currentEmail = null
        this.saveState()
        return
      }

      const itemId = item.itemId || null

      // If same item, don't re-capture
      if (itemId && itemId === this._currentItemId) return

      this._currentItemId = itemId

      const emailData = {
        itemId: itemId,
        capturedAt: new Date().toISOString()
      }

      // Capture To
      emailData.to = await this._getRecipients(item, 'to')
      // Capture From
      emailData.from = await this._getFrom(item)
      // Capture CC
      emailData.cc = await this._getRecipients(item, 'cc')
      // Capture Subject
      emailData.subject = await this._getSubject(item)
      // Capture Attachments
      emailData.attachments = this._getAttachmentNames(item)
      // Capture Internet Headers
      emailData.headers = await this._getInternetHeaders(item)

      this._state.currentEmail = emailData
      this.saveState()
      this.event('email-captured', { subject: emailData.subject, itemId: emailData.itemId })
    },

    _getRecipients(item, field) {
      return new Promise((resolve) => {
        const prop = item[field]
        if (!prop) {
          resolve([])
          return
        }
        if (typeof prop.getAsync === 'function') {
          prop.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(this._formatRecipients(result.value))
            } else {
              resolve([])
            }
          })
        } else if (Array.isArray(prop)) {
          resolve(this._formatRecipients(prop))
        } else {
          resolve([])
        }
      })
    },

    _formatRecipients(recipients) {
      if (!recipients || !Array.isArray(recipients)) return []
      return recipients.map((r) => {
        if (typeof r === 'string') return r
        return r.displayName ? r.displayName + ' <' + r.emailAddress + '>' : r.emailAddress || ''
      })
    },

    _getFrom(item) {
      return new Promise((resolve) => {
        const from = item.from
        if (!from) {
          resolve(null)
          return
        }
        if (typeof from.getAsync === 'function') {
          from.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              const val = result.value
              resolve(val.displayName ? val.displayName + ' <' + val.emailAddress + '>' : val.emailAddress || '')
            } else {
              resolve(null)
            }
          })
        } else if (typeof from === 'object' && from.emailAddress) {
          resolve(from.displayName ? from.displayName + ' <' + from.emailAddress + '>' : from.emailAddress)
        } else if (typeof from === 'string') {
          resolve(from)
        } else {
          resolve(null)
        }
      })
    },

    _getSubject(item) {
      return new Promise((resolve) => {
        const subject = item.subject
        if (subject === undefined || subject === null) {
          resolve('')
          return
        }
        if (typeof subject === 'object' && typeof subject.getAsync === 'function') {
          subject.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || '')
            } else {
              resolve('')
            }
          })
        } else {
          resolve(String(subject))
        }
      })
    },

    _getAttachmentNames(item) {
      if (!item.attachments) return []
      return item.attachments.map((a) => a.name || 'Unnamed attachment')
    },

    _getInternetHeaders(item) {
      return new Promise((resolve) => {
        if (typeof item.getAllInternetHeadersAsync === 'function') {
          item.getAllInternetHeadersAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(this._parseInternetHeaders(result.value))
            } else {
              resolve({})
            }
          })
        } else {
          resolve({})
        }
      })
    },

    _parseInternetHeaders(headerString) {
      const headers = {}
      if (!headerString || typeof headerString !== 'string') return headers

      // Unfold continuation lines (lines starting with whitespace are continuations)
      const unfolded = headerString.replace(/\r?\n[ \t]+/g, ' ')
      const lines = unfolded.split(/\r?\n/)

      lines.forEach((line) => {
        const colonIndex = line.indexOf(':')
        if (colonIndex > 0) {
          const key = line.substring(0, colonIndex).trim()
          const value = line.substring(colonIndex + 1).trim()
          if (key) {
            if (headers[key]) {
              if (Array.isArray(headers[key])) {
                headers[key].push(value)
              } else {
                headers[key] = [headers[key], value]
              }
            } else {
              headers[key] = value
            }
          }
        }
      })
      return headers
    },

    async _updateFolderFromHeaders() {
      const email = this._state.currentEmail
      if (!email || !email.headers) return

      const headers = email.headers
      let folderName = null

      // Try known headers that might contain folder information
      if (headers['X-Folder']) {
        folderName = Array.isArray(headers['X-Folder']) ? headers['X-Folder'][0] : headers['X-Folder']
      } else if (headers['X-Mailbox']) {
        folderName = Array.isArray(headers['X-Mailbox']) ? headers['X-Mailbox'][0] : headers['X-Mailbox']
      } else if (headers['X-MS-Exchange-Organization-AuthSource']) {
        const val = headers['X-MS-Exchange-Organization-AuthSource']
        folderName = Array.isArray(val) ? val[0] : val
      }

      if (folderName && folderName !== this._state.folderName) {
        this._state.folderName = folderName
        this.saveState()
      }
    },

    async notify() {
      const emailData = this._state.currentEmail
      if (!emailData) return

      const payload = {
        userName: this._state.userName,
        folderName: this._state.folderName,
        platform: this._state.platform,
        version: this._state.version,
        email: emailData
      }

      this.event('notify', { subject: emailData.subject, itemId: emailData.itemId })

      try {
        await fetch('https://jamesladd.github.io/aladdin/api/notify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        })
      } catch (e) {
        console.warn('Notify API call failed (expected for stub)', e)
      }

      // Clear current email after notify
      this._state.currentEmail = null
      this._currentItemId = null
      this.saveState()
    },

    async _addMasterCategories() {
      const categories = await this.getCategories()
      const mailbox = this.Office.context.mailbox

      if (!mailbox.masterCategories || typeof mailbox.masterCategories.addAsync !== 'function') {
        console.warn('masterCategories.addAsync not supported')
        return
      }

      const categoryList = categories.map((cat) => ({
        displayName: cat.displayName,
        color: this.Office.MailboxEnums.CategoryColor.Preset0
      }))

      return new Promise((resolve) => {
        mailbox.masterCategories.addAsync(categoryList, (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('Master categories added')
          } else {
            console.warn('Failed to add master categories', result.error)
          }
          resolve()
        })
      })
    },

    async getCategories() {
      return [
        { displayName: 'InboxAgent', color: '#f54900' }
      ]
    },

    renderUI() {
      if (typeof document === 'undefined') return

      // Render user info
      const userNameEl = document.getElementById('user-name')
      if (userNameEl) {
        userNameEl.textContent = this._state.userName || 'Unknown User'
      }

      const folderNameEl = document.getElementById('folder-name')
      if (folderNameEl) {
        folderNameEl.textContent = 'Folder: ' + (this._state.folderName || 'Inbox (default)')
      }

      const platformEl = document.getElementById('platform')
      if (platformEl) {
        platformEl.textContent = 'Platform: ' + (this._state.platform || 'Unknown')
      }

      const versionEl = document.getElementById('version')
      if (versionEl) {
        versionEl.textContent = 'Version: ' + (this._state.version || 'Unknown')
      }

      // Render current email
      const emailEl = document.getElementById('current-email')
      if (emailEl) {
        const email = this._state.currentEmail
        if (email) {
          let html = '<div class="email-detail">'
          html += '<div class="email-field"><span class="email-field-label">Subject: </span><span class="email-field-value">' + this._escapeHtml(email.subject || '(no subject)') + '</span></div>'
          html += '<div class="email-field"><span class="email-field-label">From: </span><span class="email-field-value">' + this._escapeHtml(email.from || '(unknown)') + '</span></div>'
          html += '<div class="email-field"><span class="email-field-label">To: </span><span class="email-field-value">' + this._escapeHtml((email.to || []).join(', ') || '(none)') + '</span></div>'
          html += '<div class="email-field"><span class="email-field-label">CC: </span><span class="email-field-value">' + this._escapeHtml((email.cc || []).join(', ') || '(none)') + '</span></div>'
          html += '<div class="email-field"><span class="email-field-label">Attachments: </span><span class="email-field-value">' + this._escapeHtml((email.attachments || []).join(', ') || '(none)') + '</span></div>'

          if (email.headers && Object.keys(email.headers).length > 0) {
            html += '<div class="email-field"><span class="email-field-label">Headers:</span></div>'
            html += '<div class="email-headers">'
            const headerKeys = Object.keys(email.headers)
            headerKeys.forEach((key) => {
              const val = email.headers[key]
              const display = Array.isArray(val) ? val.join('; ') : val
              html += this._escapeHtml(key) + ': ' + this._escapeHtml(display) + '\n'
            })
            html += '</div>'
          }

          html += '</div>'
          emailEl.innerHTML = html
        } else {
          emailEl.innerHTML = '<div class="no-events">No email selected</div>'
        }
      }

      // Render events
      const eventsEl = document.getElementById('events')
      if (eventsEl) {
        if (this._state.events.length === 0) {
          eventsEl.innerHTML = '<div class="no-events">No events recorded</div>'
        } else {
          let html = ''
          this._state.events.forEach((evt) => {
            html += '<div class="event-item">'
            html += '<div class="event-name">' + this._escapeHtml(evt.name) + '</div>'
            html += '<div class="event-timestamp">' + this._escapeHtml(evt.timestamp) + '</div>'
            if (evt.details && Object.keys(evt.details).length > 0) {
              html += '<div class="event-details">' + this._escapeHtml(JSON.stringify(evt.details, null, 2)) + '</div>'
            }
            html += '</div>'
          })
          eventsEl.innerHTML = html
        }
      }
    },

    _escapeHtml(str) {
      if (!str) return ''
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;')
    }
  }
}