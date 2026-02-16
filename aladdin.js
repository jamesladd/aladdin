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
    _state: {
      events: [],
      userName: null,
      folderName: null,
      platform: null,
      version: null,
      currentEmail: null
    },

    state() {
      return this._state
    },

    saveState() {
      try {
        if (typeof localStorage !== 'undefined') {
          localStorage.setItem('aladdin_state', JSON.stringify(this._state))
        }
      } catch (e) {
        console.error('Failed to save state', e)
      }
    },

    loadState() {
      try {
        if (typeof localStorage !== 'undefined') {
          const saved = localStorage.getItem('aladdin_state')
          if (saved) {
            const parsed = JSON.parse(saved)
            this._state = parsed
            if (!this._state.events) this._state.events = []
          }
        }
      } catch (e) {
        console.error('Failed to load state', e)
      }
    },

    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === 'aladdin_state' && e.newValue) {
            try {
              this._state = JSON.parse(e.newValue)
              if (!this._state.events) this._state.events = []
              this.updateUI()
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
        name: name,
        details: details,
        timestamp: new Date().toISOString()
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

      // Get user name
      const userName = mailbox.userProfile ? mailbox.userProfile.displayName : 'Unknown User'
      this._state.userName = userName

      // Get platform
      const platform = this.Office.context.platform || 'Unknown'
      this._state.platform = platform

      // Get version
      let version = 'Unknown'
      if (this.Office.context.diagnostics && this.Office.context.diagnostics.version) {
        version = this.Office.context.diagnostics.version
      }
      this._state.version = version

      // Get folder name - attempt from current item headers, fallback to default
      this._state.folderName = 'Inbox (default)'
      this.saveState()

      // Register mailbox-level events
      this.registerMailboxEvents()

      // Register item-level events and capture current item
      if (mailbox.item) {
        this.registerItemEvents(mailbox.item)
        this.captureCurrentItem()
      }

      // Add master categories
      this.addMasterCategories()

      this.event('initialized', {
        userName: userName,
        folderName: this._state.folderName,
        platform: platform,
        version: version
      })

      this.updateUI()
    },

    registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox
      const self = this

      // ItemChanged - mailbox level event (fires when selected item changes)
      if (this.Office.EventType.ItemChanged) {
        try {
          mailbox.addHandlerAsync(this.Office.EventType.ItemChanged, function (eventArgs) {
            self.handleItemChanged(eventArgs)
          })
        } catch (e) {
          console.error('Failed to register ItemChanged', e)
        }
      }

      // OfficeThemeChanged - mailbox level event
      if (this.Office.EventType.OfficeThemeChanged) {
        try {
          mailbox.addHandlerAsync(this.Office.EventType.OfficeThemeChanged, function (eventArgs) {
            self.event('OfficeThemeChanged', { type: 'OfficeThemeChanged' })
          })
        } catch (e) {
          console.error('Failed to register OfficeThemeChanged', e)
        }
      }
    },

    registerItemEvents(item) {
      if (!item) return
      const self = this

      // RecipientsChanged - item level event
      if (this.Office.EventType.RecipientsChanged) {
        try {
          item.addHandlerAsync(this.Office.EventType.RecipientsChanged, function (eventArgs) {
            self.event('RecipientsChanged', { type: 'RecipientsChanged' })
            self.captureCurrentItem()
          })
        } catch (e) {
          console.error('Failed to register RecipientsChanged', e)
        }
      }

      // AttachmentsChanged - item level event
      if (this.Office.EventType.AttachmentsChanged) {
        try {
          item.addHandlerAsync(this.Office.EventType.AttachmentsChanged, function (eventArgs) {
            self.event('AttachmentsChanged', { type: 'AttachmentsChanged' })
            self.captureCurrentItem()
          })
        } catch (e) {
          console.error('Failed to register AttachmentsChanged', e)
        }
      }

      // RecurrenceChanged - item level event
      if (this.Office.EventType.RecurrenceChanged) {
        try {
          item.addHandlerAsync(this.Office.EventType.RecurrenceChanged, function (eventArgs) {
            self.event('RecurrenceChanged', { type: 'RecurrenceChanged' })
          })
        } catch (e) {
          console.error('Failed to register RecurrenceChanged', e)
        }
      }

      // AppointmentTimeChanged - item level event
      if (this.Office.EventType.AppointmentTimeChanged) {
        try {
          item.addHandlerAsync(this.Office.EventType.AppointmentTimeChanged, function (eventArgs) {
            self.event('AppointmentTimeChanged', { type: 'AppointmentTimeChanged' })
          })
        } catch (e) {
          console.error('Failed to register AppointmentTimeChanged', e)
        }
      }

      // EnhancedLocationsChanged - item level event
      if (this.Office.EventType.EnhancedLocationsChanged) {
        try {
          item.addHandlerAsync(this.Office.EventType.EnhancedLocationsChanged, function (eventArgs) {
            self.event('EnhancedLocationsChanged', { type: 'EnhancedLocationsChanged' })
          })
        } catch (e) {
          console.error('Failed to register EnhancedLocationsChanged', e)
        }
      }
    },

    handleItemChanged(eventArgs) {
      const self = this

      // Notify about previous item before switching
      self.notifyPreviousItem()

      const mailbox = self.Office.context.mailbox

      if (mailbox.item) {
        self.registerItemEvents(mailbox.item)
        self.captureCurrentItem()
        self.event('ItemChanged', { type: 'ItemChanged', hasItem: true })
      } else {
        self._currentItemId = null
        self._state.currentEmail = null
        self.saveState()
        self.event('ItemChanged', { type: 'ItemChanged', hasItem: false })
      }

      self.updateUI()
    },

    notifyPreviousItem() {
      if (this._state.currentEmail) {
        this.notify(this._state.currentEmail)
        this._state.currentEmail = null
        this.saveState()
      }
    },

    async captureCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const emailData = {
        itemId: item.itemId || null,
        to: null,
        from: null,
        cc: null,
        subject: null,
        attachments: [],
        headers: {},
        capturedAt: new Date().toISOString()
      }

      // Get the item ID for tracking
      this._currentItemId = item.itemId || null

      try {
        // Get To - check for getAsync (compose mode) vs synchronous (read mode)
        emailData.to = await this._getField(item, 'to')

        // Get From - check for getAsync vs synchronous
        emailData.from = await this._getField(item, 'from')

        // Get Subject - check for getAsync vs synchronous
        emailData.subject = await this._getFieldValue(item, 'subject')

        // Get CC - check for getAsync vs synchronous
        emailData.cc = await this._getField(item, 'cc')

        // Get Attachments
        if (item.attachments) {
          emailData.attachments = item.attachments.map(function (att) {
            return att.name
          })
        }

        // Get Internet Headers
        emailData.headers = await this._getInternetHeaders(item)

        // Try to extract folder info from headers
        this._extractFolderFromHeaders(emailData.headers)

      } catch (e) {
        console.error('Error capturing item', e)
      }

      this._state.currentEmail = emailData
      this.saveState()
      this.updateUI()
    },

    _getField(item, fieldName) {
      return new Promise(function (resolve) {
        const field = item[fieldName]
        if (!field) {
          resolve(null)
          return
        }
        // Check if the field has getAsync (compose mode)
        if (field && typeof field.getAsync === 'function') {
          field.getAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(formatRecipients(result.value))
            } else {
              resolve(null)
            }
          })
        } else {
          // Synchronous access (read mode)
          resolve(formatRecipients(field))
        }
      })

      function formatRecipients(value) {
        if (!value) return null
        if (Array.isArray(value)) {
          return value.map(function (r) {
            return r.displayName || r.emailAddress || String(r)
          }).join(', ')
        }
        if (typeof value === 'object') {
          return value.displayName || value.emailAddress || JSON.stringify(value)
        }
        return String(value)
      }
    },

    _getFieldValue(item, fieldName) {
      const self = this
      return new Promise(function (resolve) {
        const field = item[fieldName]
        // Check if the field itself has getAsync (compose mode for subject)
        if (field && typeof field === 'object' && typeof field.getAsync === 'function') {
          field.getAsync(function (result) {
            if (result.status === self.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value)
            } else {
              resolve(null)
            }
          })
        } else {
          // Synchronous access (read mode) - subject is a string directly
          resolve(field || null)
        }
      })
    },

    _getInternetHeaders(item) {
      const self = this
      return new Promise(function (resolve) {
        if (item.getAllInternetHeadersAsync && typeof item.getAllInternetHeadersAsync === 'function') {
          item.getAllInternetHeadersAsync(function (result) {
            if (result.status === self.Office.AsyncResultStatus.Succeeded && result.value) {
              resolve(self._parseHeaders(result.value))
            } else {
              resolve({})
            }
          })
        } else {
          resolve({})
        }
      })
    },

    _parseHeaders(headerString) {
      const headers = {}
      if (!headerString || typeof headerString !== 'string') return headers

      // Split on line breaks, handle continuation lines (lines starting with whitespace)
      const lines = headerString.split(/\r?\n/)
      let currentKey = null
      let currentValue = ''

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i]
        if (!line) continue

        // Continuation line (starts with space or tab)
        if (line.charAt(0) === ' ' || line.charAt(0) === '\t') {
          if (currentKey) {
            currentValue += ' ' + line.trim()
          }
        } else {
          // Save previous header
          if (currentKey) {
            headers[currentKey] = currentValue
          }
          // Parse new header
          const colonIndex = line.indexOf(':')
          if (colonIndex > 0) {
            currentKey = line.substring(0, colonIndex).trim()
            currentValue = line.substring(colonIndex + 1).trim()
          } else {
            currentKey = null
            currentValue = ''
          }
        }
      }

      // Save last header
      if (currentKey) {
        headers[currentKey] = currentValue
      }

      return headers
    },

    _extractFolderFromHeaders(headers) {
      if (!headers) return

      // Try known headers that may contain folder information
      const folderHeaders = [
        'X-Folder',
        'X-MS-Exchange-Organization-AuthSource',
        'X-MS-Exchange-Parent-Message-Id',
        'X-Mailer'
      ]

      for (let i = 0; i < folderHeaders.length; i++) {
        if (headers[folderHeaders[i]]) {
          this._state.folderName = headers[folderHeaders[i]]
          this.saveState()
          return
        }
      }
    },

    notify(emailData) {
      console.log('Notify API - email deselected:', emailData)
      // Future: make actual API call here
    },

    async getCategories(userInfo) {
      return [
        { displayName: 'InboxAgent', color: 'Preset0' }
      ]
    },

    async addMasterCategories() {
      const self = this
      const userInfo = {
        userName: this._state.userName,
        folderName: this._state.folderName,
        platform: this._state.platform,
        version: this._state.version
      }

      try {
        const categories = await this.getCategories(userInfo)
        const masterCategories = this.Office.context.mailbox.masterCategories

        if (!masterCategories || typeof masterCategories.addAsync !== 'function') {
          console.log('masterCategories.addAsync not available')
          return
        }

        masterCategories.addAsync(categories, function (result) {
          if (result.status === self.Office.AsyncResultStatus.Succeeded) {
            console.log('Master categories added successfully')
          } else {
            console.log('Failed to add master categories:', result.error ? result.error.message : 'unknown error')
          }
        })
      } catch (e) {
        console.error('Error adding master categories', e)
      }
    },

    updateUI() {
      if (typeof document === 'undefined') return

      // Update user name
      const userNameEl = document.getElementById('user-name')
      if (userNameEl) {
        userNameEl.textContent = this._state.userName || 'Unknown User'
      }

      // Update folder name
      const folderNameEl = document.getElementById('folder-name')
      if (folderNameEl) {
        folderNameEl.textContent = this._state.folderName || 'Unknown'
      }

      // Update platform
      const platformEl = document.getElementById('platform')
      if (platformEl) {
        platformEl.textContent = this._state.platform || 'Unknown'
      }

      // Update version
      const versionEl = document.getElementById('version')
      if (versionEl) {
        versionEl.textContent = this._state.version || 'Unknown'
      }

      // Update status
      const statusEl = document.getElementById('status')
      if (statusEl) {
        statusEl.textContent = 'Active'
        statusEl.className = 'status-active'
      }

      // Update current email
      const currentEmailEl = document.getElementById('current-email')
      if (currentEmailEl) {
        const email = this._state.currentEmail
        if (email) {
          let html = ''
          html += '<div class="email-detail"><span class="info-label">Subject:</span> <span class="info-value">' + escapeHtml(email.subject || '(none)') + '</span></div>'
          html += '<div class="email-detail"><span class="info-label">From:</span> <span class="info-value">' + escapeHtml(email.from || '(none)') + '</span></div>'
          html += '<div class="email-detail"><span class="info-label">To:</span> <span class="info-value">' + escapeHtml(email.to || '(none)') + '</span></div>'
          html += '<div class="email-detail"><span class="info-label">CC:</span> <span class="info-value">' + escapeHtml(email.cc || '(none)') + '</span></div>'

          if (email.attachments && email.attachments.length > 0) {
            html += '<div class="email-detail"><span class="info-label">Attachments:</span> <span class="info-value">' + escapeHtml(email.attachments.join(', ')) + '</span></div>'
          } else {
            html += '<div class="email-detail"><span class="info-label">Attachments:</span> <span class="info-value">(none)</span></div>'
          }

          if (email.headers && Object.keys(email.headers).length > 0) {
            html += '<div class="email-detail"><span class="info-label">Headers:</span></div>'
            html += '<div class="headers-container">'
            const keys = Object.keys(email.headers)
            for (let i = 0; i < keys.length; i++) {
              html += '<div class="header-item"><span class="header-key">' + escapeHtml(keys[i]) + ':</span> <span class="header-value">' + escapeHtml(email.headers[keys[i]]) + '</span></div>'
            }
            html += '</div>'
          }

          currentEmailEl.innerHTML = html
        } else {
          currentEmailEl.innerHTML = '<div class="no-email">No email selected</div>'
        }
      }

      // Update events
      const eventsEl = document.getElementById('events')
      if (eventsEl) {
        const events = this._state.events || []
        if (events.length === 0) {
          eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        } else {
          let html = ''
          // Show most recent first
          for (let i = events.length - 1; i >= 0; i--) {
            const ev = events[i]
            html += '<div class="event-item">'
            html += '<div class="event-name">' + escapeHtml(ev.name) + '</div>'
            html += '<div class="event-timestamp">' + escapeHtml(ev.timestamp) + '</div>'
            if (ev.details) {
              html += '<div class="event-details">' + escapeHtml(JSON.stringify(ev.details, null, 2)) + '</div>'
            }
            html += '</div>'
          }
          eventsEl.innerHTML = html
        }
      }
    }
  }
}

function escapeHtml(str) {
  if (!str) return ''
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}