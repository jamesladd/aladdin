// aladdin.js

const singleton = [false]

export function createAladdin(Office) {
  if (typeof window !== 'undefined' && window.aladdinInstance) return window.aladdinInstance;
  if (singleton[0]) return singleton[0];
  const instance = aladdin(Office)
  if (typeof window !== 'undefined') window.aladdinInstance = instance;
  if (typeof window === 'undefined') singleton[0] = instance;
  instance.loadState()
  instance.watchState()
  return instance
}

function aladdin(Office) {
  return {
    Office,
    _currentItemId: null,
    _itemHandlersRegistered: false,
    _state: {
      events: [],
      currentEmail: null,
      userInfo: null,
    },
    state() {
      return this._state
    },
    saveState() {
      try {
        localStorage.setItem('aladdin_state', JSON.stringify(this._state))
      } catch (e) {
        console.error('Failed to save state', e)
      }
    },
    loadState() {
      try {
        const saved = localStorage.getItem('aladdin_state')
        if (saved) {
          const parsed = JSON.parse(saved)
          this._state = parsed
          if (!this._state.events) this._state.events = []
          if (this._state.currentEmail) {
            this._currentItemId = this._state.currentEmail.itemId || null
          }
        }
      } catch (e) {
        console.error('Failed to load state', e)
      }
      this.updateUI()
    },
    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === 'aladdin_state') {
            try {
              const parsed = JSON.parse(e.newValue)
              if (parsed) {
                this._state = parsed
                if (!this._state.events) this._state.events = []
                this.updateUI()
              }
            } catch (err) {
              console.error('Failed to parse storage event', err)
            }
          }
        })
      }
    },
    event(name, details) {
      console.log('Event:', name)
      const entry = {
        name,
        details: details || {},
        timestamp: new Date().toISOString(),
      }
      this._state.events.push(entry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(this._state.events.length - 10)
      }
      this.saveState()
      this.updateUI()
    },
    initialize() {
      const mailbox = this.Office.context.mailbox
      const userName = mailbox.userProfile.displayName || 'Unknown User'
      const platform = this.Office.context.platform || 'Unknown'
      let version = 'Unknown'
      try {
        version = this.Office.context.diagnostics.version || 'Unknown'
      } catch (e) {
        version = 'Unknown'
      }

      this._state.userInfo = {
        userName,
        folder: 'Inbox (default)',
        platform,
        version,
      }
      this.saveState()
      this.updateUI()

      this._registerMailboxEvents()

      if (mailbox.item) {
        this._captureCurrentItem(() => {
          this._tryExtractFolder()
          this._registerItemEvents()
          this._initCategories()
        })
      } else {
        this._initCategories()
      }
    },
    _initCategories() {
      const info = this._state.userInfo || {}
      this.getCategories(info).then((categories) => {
        this._addMasterCategories(categories)
      })
    },
    _addMasterCategories(categories) {
      const mailbox = this.Office.context.mailbox
      if (mailbox.masterCategories && mailbox.masterCategories.addAsync) {
        const catList = categories.map((c) => ({
          displayName: c.displayName,
          color: this.Office.MailboxEnums.CategoryColor.Preset0,
        }))
        mailbox.masterCategories.addAsync(catList, (result) => {
          if (result.status === this.Office.AsyncResultStatus.Failed) {
            console.error('Failed to add master categories', result.error)
          }
        })
      }
    },
    _registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox
      const EventType = this.Office.EventType

      if (EventType.ItemChanged && mailbox.addHandlerAsync) {
        mailbox.addHandlerAsync(EventType.ItemChanged, (eventArgs) => {
          this.event('ItemChanged', { type: 'ItemChanged' })
          this._onItemChanged()
        })
      }

      if (EventType.OfficeThemeChanged && mailbox.addHandlerAsync) {
        mailbox.addHandlerAsync(EventType.OfficeThemeChanged, (eventArgs) => {
          this.event('OfficeThemeChanged', { type: 'OfficeThemeChanged' })
        })
      }
    },
    _registerItemEvents() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.addHandlerAsync) {
        this._itemHandlersRegistered = false
        return
      }
      if (this._itemHandlersRegistered) return
      this._itemHandlersRegistered = true

      const EventType = this.Office.EventType
      const itemEvents = [
        'RecipientsChanged',
        'AttachmentsChanged',
        'RecurrenceChanged',
        'AppointmentTimeChanged',
        'EnhancedLocationsChanged',
      ]

      itemEvents.forEach((evtName) => {
        if (EventType[evtName]) {
          item.addHandlerAsync(EventType[evtName], (eventArgs) => {
            this.event(evtName, { type: evtName })
            if (evtName === 'RecipientsChanged' || evtName === 'AttachmentsChanged') {
              this._captureCurrentItem()
            }
          })
        }
      })
    },
    _onItemChanged() {
      const previousEmail = this._state.currentEmail
      if (previousEmail) {
        this.notify(previousEmail)
        this._state.currentEmail = null
        this._currentItemId = null
        this.saveState()
      }

      this._itemHandlersRegistered = false

      const mailbox = this.Office.context.mailbox
      if (mailbox.item) {
        this._captureCurrentItem(() => {
          this._tryExtractFolder()
          this._registerItemEvents()
        })
      } else {
        this.updateUI()
      }
    },
    _captureCurrentItem(callback) {
      const item = this.Office.context.mailbox.item
      if (!item) {
        if (callback) callback()
        return
      }

      const email = {
        itemId: item.itemId || null,
        to: [],
        from: null,
        cc: [],
        subject: '',
        attachments: [],
        headers: {},
        categories: [],
        capturedAt: new Date().toISOString(),
      }

      let pending = 0
      const done = () => {
        pending--
        if (pending <= 0) {
          this._state.currentEmail = email
          this._currentItemId = email.itemId
          this.saveState()
          this.updateUI()
          if (callback) callback()
        }
      }

      // To
      pending++
      if (item.to && typeof item.to.getAsync === 'function') {
        item.to.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            email.to = (result.value || []).map((r) => ({ name: r.displayName, email: r.emailAddress }))
          }
          done()
        })
      } else if (item.to) {
        email.to = (item.to || []).map((r) => ({ name: r.displayName, email: r.emailAddress }))
        done()
      } else {
        done()
      }

      // From
      pending++
      if (item.from && typeof item.from.getAsync === 'function') {
        item.from.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded && result.value) {
            email.from = { name: result.value.displayName, email: result.value.emailAddress }
          }
          done()
        })
      } else if (item.from) {
        email.from = { name: item.from.displayName, email: item.from.emailAddress }
        done()
      } else if (item.sender) {
        email.from = { name: item.sender.displayName, email: item.sender.emailAddress }
        done()
      } else {
        done()
      }

      // CC
      pending++
      if (item.cc && typeof item.cc.getAsync === 'function') {
        item.cc.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            email.cc = (result.value || []).map((r) => ({ name: r.displayName, email: r.emailAddress }))
          }
          done()
        })
      } else if (item.cc) {
        email.cc = (item.cc || []).map((r) => ({ name: r.displayName, email: r.emailAddress }))
        done()
      } else {
        done()
      }

      // Subject
      pending++
      if (item.subject && typeof item.subject.getAsync === 'function') {
        item.subject.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            email.subject = result.value || ''
          }
          done()
        })
      } else if (typeof item.subject === 'string') {
        email.subject = item.subject
        done()
      } else {
        done()
      }

      // Attachments
      pending++
      if (item.attachments) {
        email.attachments = (item.attachments || []).map((a) => a.name)
      }
      done()

      // Categories
      pending++
      if (item.categories && typeof item.categories.getAsync === 'function') {
        item.categories.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            email.categories = (result.value || []).map((c) => c.displayName)
          }
          done()
        })
      } else if (item.categories && Array.isArray(item.categories)) {
        email.categories = item.categories.map((c) => (typeof c === 'string' ? c : c.displayName))
        done()
      } else {
        done()
      }

      // Internet Headers
      pending++
      if (item.getAllInternetHeadersAsync) {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded && result.value) {
            email.headers = this._parseHeaders(result.value)
          }
          done()
        })
      } else {
        done()
      }
    },
    _parseHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers
      const unfolded = headerString.replace(/\r?\n[ \t]+/g, ' ')
      const lines = unfolded.split(/\r?\n/)
      lines.forEach((line) => {
        const idx = line.indexOf(':')
        if (idx > 0) {
          const key = line.substring(0, idx).trim()
          const value = line.substring(idx + 1).trim()
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
      })
      return headers
    },
    _tryExtractFolder() {
      if (!this._state.currentEmail || !this._state.currentEmail.headers) return
      const headers = this._state.currentEmail.headers
      const folderKeys = [
        'X-Folder',
        'X-MS-Exchange-Organization-FolderName',
        'X-Folder-Name',
        'X-GM-LABELS',
        'X-Microsoft-Antispam-Mailbox-Delivery',
      ]
      for (const key of folderKeys) {
        if (headers[key]) {
          const val = Array.isArray(headers[key]) ? headers[key][0] : headers[key]
          if (val) {
            this._state.userInfo.folder = val
            this.saveState()
            this.updateUI()
            return
          }
        }
      }
    },
    async notify(emailData) {
      if (!emailData) return
      this.event('notify', { subject: emailData.subject })
      try {
        await fetch('/api/inboxnotify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(emailData),
        })
      } catch (e) {
        console.error('Failed to notify API', e)
      }
    },
    async getCategories(info) {
      try {
        const response = await fetch('/api/inboxinit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(info || {}),
        })
        if (response.ok) {
          const data = await response.json()
          if (Array.isArray(data) && data.length > 0) {
            return data
          }
        }
      } catch (e) {
        console.error('Failed to get categories from API, using defaults', e)
      }
      return [{ displayName: 'InboxAgent', color: '#f54900' }]
    },
    updateUI() {
      const statusEl = document.getElementById('status')
      const eventsEl = document.getElementById('events')
      const userInfoEl = document.getElementById('user-info')
      const emailDetailEl = document.getElementById('email-detail')

      if (!statusEl) return

      // User info
      if (userInfoEl && this._state.userInfo) {
        const info = this._state.userInfo
        userInfoEl.innerHTML =
          '<div class="user-info">' +
          '<div class="user-name">' + this._escapeHtml(info.userName) + '</div>' +
          '<div class="platform">Folder: ' + this._escapeHtml(info.folder) + '</div>' +
          '<div class="platform">Platform: ' + this._escapeHtml(info.platform) + '</div>' +
          '<div class="platform">Version: ' + this._escapeHtml(info.version) + '</div>' +
          '</div>'
      }

      // Current email
      if (emailDetailEl) {
        const email = this._state.currentEmail
        if (email) {
          const fromStr = email.from ? (email.from.name + ' <' + email.from.email + '>') : 'N/A'
          const toStr = (email.to || []).map((r) => r.name + ' <' + r.email + '>').join(', ') || 'N/A'
          const ccStr = (email.cc || []).map((r) => r.name + ' <' + r.email + '>').join(', ') || 'N/A'
          const attachStr = (email.attachments || []).join(', ') || 'None'
          const catStr = (email.categories || []).join(', ') || 'None'
          const headerCount = Object.keys(email.headers || {}).length

          emailDetailEl.innerHTML =
            '<div class="section-title">Current Email</div>' +
            '<div class="email-detail-card">' +
            '<div class="email-field"><strong>Subject:</strong> ' + this._escapeHtml(email.subject || 'N/A') + '</div>' +
            '<div class="email-field"><strong>From:</strong> ' + this._escapeHtml(fromStr) + '</div>' +
            '<div class="email-field"><strong>To:</strong> ' + this._escapeHtml(toStr) + '</div>' +
            '<div class="email-field"><strong>CC:</strong> ' + this._escapeHtml(ccStr) + '</div>' +
            '<div class="email-field"><strong>Attachments:</strong> ' + this._escapeHtml(attachStr) + '</div>' +
            '<div class="email-field"><strong>Categories:</strong> ' + this._escapeHtml(catStr) + '</div>' +
            '<div class="email-field"><strong>Headers:</strong> ' + headerCount + ' header(s)</div>' +
            '</div>'
        } else {
          emailDetailEl.innerHTML =
            '<div class="section-title">Current Email</div>' +
            '<div class="no-events">No email selected</div>'
        }
      }

      // Status
      statusEl.textContent = this._state.currentEmail ? 'Email selected' : 'Ready'

      // Events
      if (eventsEl) {
        const events = this._state.events || []
        if (events.length === 0) {
          eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        } else {
          let html = ''
          const reversed = events.slice().reverse()
          reversed.forEach((evt) => {
            html +=
              '<div class="event-item">' +
              '<div class="event-name">' + this._escapeHtml(evt.name) + '</div>' +
              '<div class="event-timestamp">' + this._escapeHtml(evt.timestamp) + '</div>' +
              '<div class="event-details">' + this._escapeHtml(JSON.stringify(evt.details, null, 2)) + '</div>' +
              '</div>'
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
    },
  }
}