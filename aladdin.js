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
    _itemEventsRegistered: false,
    _state: {
      events: [],
      capturedEmail: null,
      userInfo: {
        userName: '',
        folderName: '',
        platform: '',
        version: ''
      }
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
            if (!this._state.userInfo) this._state.userInfo = { userName: '', folderName: '', platform: '', version: '' }
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
              const parsed = JSON.parse(e.newValue)
              this._state = parsed
              if (!this._state.events) this._state.events = []
              if (!this._state.userInfo) this._state.userInfo = { userName: '', folderName: '', platform: '', version: '' }
              this._currentItemId = this._state.capturedEmail ? this._state.capturedEmail.graphMessageId : null
              this._updateUI()
            } catch (e2) {
              console.error('Failed to parse storage event', e2)
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
        timestamp: new Date().toISOString()
      }
      this._state.events.push(entry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(-10)
      }
      this.saveState()
      this._updateUI()
    },

    async initialize() {
      const mailbox = this.Office.context.mailbox
      const userProfile = mailbox.userProfile

      // Gather user info (requirements 2-5)
      const userName = userProfile ? (userProfile.displayName || userProfile.emailAddress || '') : ''
      const platform = this.Office.context.platform || ''
      let version = ''
      try {
        version = this.Office.context.diagnostics ? this.Office.context.diagnostics.version : ''
      } catch (e) {
        version = ''
      }

      this._state.userInfo.userName = userName
      this._state.userInfo.platform = platform
      this._state.userInfo.version = version

      // Attempt to get folder name from headers
      await this._extractFolderName()

      this.saveState()
      this._updateUI()

      // Rule R3: On initialization, check for previously captured email that needs notification
      this.loadState()
      const item = mailbox.item
      const isCompose = item && !item.itemId
      const currentId = (item && item.itemId) ? item.itemId : null

      if (this._state.capturedEmail) {
        const previousId = this._state.capturedEmail.graphMessageId
        if (!item || isCompose || (currentId && previousId && currentId !== this._state.capturedEmail._rawItemId)) {
          await this.notify(this._state.capturedEmail)
          this._state.capturedEmail = null
          this._currentItemId = null
          this.saveState()
          this._updateUI()
        }
      }

      // Register mailbox-level events
      this._registerMailboxEvents()

      // Register item-level events if item exists
      if (item) {
        this._registerItemEvents()
      }

      // Get categories and ensure master categories
      await this._initCategories()

      // Capture current item if in read mode
      if (item && !isCompose) {
        await this._captureCurrentItem()
      }

      this.event('initialized', { userName, platform, version, folderName: this._state.userInfo.folderName })
    },

    async notify(emailData) {
      if (!emailData) return
      this.event('notify', { subject: emailData.subject, graphMessageId: emailData.graphMessageId })
      try {
        await fetch('/api/inboxnotify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(emailData)
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
          body: JSON.stringify(info)
        })
        if (response.ok) {
          const data = await response.json()
          if (Array.isArray(data) && data.length > 0) return data
        }
      } catch (e) {
        console.error('Failed to get categories from API', e)
      }
      return [
        { displayName: 'InboxAgent', color: '#f54900' }
      ]
    },

    async _initCategories() {
      const info = {
        userName: this._state.userInfo.userName,
        folderName: this._state.userInfo.folderName,
        platform: this._state.userInfo.platform,
        version: this._state.userInfo.version
      }
      const categories = await this.getCategories(info)
      await this._ensureMasterCategories(categories)
    },

    async _ensureMasterCategories(categories) {
      const mailbox = this.Office.context.mailbox
      if (!mailbox.masterCategories || !mailbox.masterCategories.addAsync) return

      for (const cat of categories) {
        try {
          await new Promise((resolve) => {
            const masterCat = {
              displayName: cat.displayName,
              color: this.Office.MailboxEnums.CategoryColor.Preset0
            }
            mailbox.masterCategories.addAsync([masterCat], (result) => {
              if (result.status === this.Office.AsyncResultStatus.Failed) {
                console.log('Category may already exist:', cat.displayName, result.error)
              }
              resolve()
            })
          })
        } catch (e) {
          console.error('Failed to add master category', cat.displayName, e)
        }
      }
    },

    async _extractFolderName() {
      const item = this.Office.context.mailbox.item
      if (!item) {
        this._state.userInfo.folderName = 'Unknown'
        return
      }

      if (item.getAllInternetHeadersAsync) {
        try {
          const headers = await new Promise((resolve) => {
            item.getAllInternetHeadersAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve(result.value)
              } else {
                resolve(null)
              }
            })
          })
          if (headers) {
            const parsed = this._parseHeaders(headers)
            const folderHeader = parsed['X-Folder'] || parsed['x-folder']
              || parsed['X-MS-Exchange-Organization-FolderName']
              || parsed['x-ms-exchange-organization-foldername']
            if (folderHeader) {
              this._state.userInfo.folderName = folderHeader
              return
            }
          }
        } catch (e) {
          console.error('Failed to get internet headers for folder', e)
        }
      }

      this._state.userInfo.folderName = 'Inbox'
    },

    _registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox
      if (!mailbox.addHandlerAsync) return

      try {
        if (this.Office.EventType.ItemChanged) {
          mailbox.addHandlerAsync(this.Office.EventType.ItemChanged, (eventArgs) => {
            this._handleItemChanged(eventArgs)
          })
        }
      } catch (e) {
        console.error('Failed to register ItemChanged', e)
      }

      try {
        if (this.Office.EventType.OfficeThemeChanged) {
          mailbox.addHandlerAsync(this.Office.EventType.OfficeThemeChanged, (eventArgs) => {
            this.event('OfficeThemeChanged', eventArgs || {})
          })
        }
      } catch (e) {
        console.log('OfficeThemeChanged not supported', e)
      }
    },

    _registerItemEvents() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.addHandlerAsync) {
        this._itemEventsRegistered = false
        return
      }

      this._itemEventsRegistered = true

      const itemEvents = [
        { type: 'RecipientsChanged', eventType: this.Office.EventType.RecipientsChanged },
        { type: 'AttachmentsChanged', eventType: this.Office.EventType.AttachmentsChanged },
        { type: 'RecurrenceChanged', eventType: this.Office.EventType.RecurrenceChanged },
        { type: 'AppointmentTimeChanged', eventType: this.Office.EventType.AppointmentTimeChanged },
        { type: 'EnhancedLocationsChanged', eventType: this.Office.EventType.EnhancedLocationsChanged }
      ]

      for (const evt of itemEvents) {
        if (evt.eventType) {
          try {
            item.addHandlerAsync(evt.eventType, (eventArgs) => {
              this.event(evt.type, eventArgs || {})
            })
          } catch (e) {
            console.log(evt.type + ' not supported', e)
          }
        }
      }
    },

    async _handleItemChanged(eventArgs) {
      this.event('ItemChanged', eventArgs || {})

      // Rule R2: Re-read state from localStorage before notifying
      this.loadState()

      const mailbox = this.Office.context.mailbox
      const item = mailbox.item
      const isCompose = item && !item.itemId
      const newItemId = (item && item.itemId) ? item.itemId : null

      // Determine if previous email needs notification
      if (this._state.capturedEmail) {
        const prevRawId = this._state.capturedEmail._rawItemId
        const isDifferent = !item || isCompose || (newItemId && newItemId !== prevRawId) || (!newItemId && !isCompose)

        if (isDifferent) {
          await this.notify(this._state.capturedEmail)
          this._state.capturedEmail = null
          this._currentItemId = null
          this.saveState()
          this._updateUI()
        }
      }

      // Re-register item-level events for the new item
      this._itemEventsRegistered = false
      if (item) {
        this._registerItemEvents()
      }

      // Rule R1: Compose mode is explicit deselection, do not capture
      if (isCompose) {
        this._currentItemId = null
        await this._extractFolderName()
        this.saveState()
        this._updateUI()
        return
      }

      // No item
      if (!item) {
        this._currentItemId = null
        this.saveState()
        this._updateUI()
        return
      }

      // New read item — capture it
      await this._extractFolderName()
      await this._captureCurrentItem()
    },

    async _captureCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const emailData = {}

      // Store raw item ID for comparison
      emailData._rawItemId = item.itemId || null

      // Graph-compatible message ID
      if (item.itemId && this.Office.context.mailbox.convertToRestId) {
        try {
          emailData.graphMessageId = this.Office.context.mailbox.convertToRestId(
            item.itemId,
            this.Office.MailboxEnums.RestVersion.v2_0
          )
        } catch (e) {
          emailData.graphMessageId = item.itemId
        }
      } else {
        emailData.graphMessageId = item.itemId || null
      }

      // To
      emailData.to = await this._getFieldAsync(item, 'to')

      // From
      emailData.from = await this._getFieldAsync(item, 'from')

      // CC
      emailData.cc = await this._getFieldAsync(item, 'cc')

      // Subject
      emailData.subject = await this._getFieldAsync(item, 'subject')

      // Categories
      emailData.categories = await this._getFieldAsync(item, 'categories')

      // Importance
      emailData.importance = item.importance || null

      // Sentiment (not natively available)
      emailData.sentiment = null

      // Attachments
      if (item.attachments) {
        emailData.attachments = item.attachments.map((a) => a.name)
      } else {
        emailData.attachments = []
      }

      // Raw headers
      emailData.headers = {}
      if (item.getAllInternetHeadersAsync) {
        try {
          const headersStr = await new Promise((resolve) => {
            item.getAllInternetHeadersAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve(result.value)
              } else {
                resolve(null)
              }
            })
          })
          if (headersStr) {
            emailData.headers = this._parseHeaders(headersStr)
          }
        } catch (e) {
          console.error('Failed to get internet headers', e)
        }
      }

      this._currentItemId = emailData._rawItemId
      this._state.capturedEmail = emailData
      this.saveState()
      this._updateUI()

      this.event('emailCaptured', { subject: emailData.subject, graphMessageId: emailData.graphMessageId })
    },

    async _getFieldAsync(item, fieldName) {
      // Requirement 15: Check for getAsync before using async methods
      // Fields that have async getters: to, from, subject, cc, categories
      const asyncMethodMap = {
        to: 'to',
        from: 'from',
        cc: 'cc',
        subject: 'subject',
        categories: 'categories'
      }

      const prop = asyncMethodMap[fieldName]
      if (!prop) return null

      // Check if the property object has getAsync
      const propObj = item[prop]

      if (propObj && typeof propObj === 'object' && typeof propObj.getAsync === 'function') {
        try {
          return await new Promise((resolve) => {
            propObj.getAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve(result.value)
              } else {
                resolve(null)
              }
            })
          })
        } catch (e) {
          console.error('Failed async get for ' + fieldName, e)
          return null
        }
      }

      // Synchronous fallback
      if (propObj !== undefined) {
        // Format recipients for to, from, cc
        if (fieldName === 'to' || fieldName === 'cc') {
          if (Array.isArray(propObj)) {
            return propObj.map((r) => ({ displayName: r.displayName, emailAddress: r.emailAddress }))
          }
          return propObj
        }
        if (fieldName === 'from') {
          if (propObj && propObj.emailAddress) {
            return { displayName: propObj.displayName, emailAddress: propObj.emailAddress }
          }
          return propObj
        }
        return propObj
      }

      return null
    },

    _parseHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers

      // Headers are in the format "Key: Value\r\n", with possible continuation lines starting with whitespace
      const lines = headerString.split(/\r?\n/)
      let currentKey = null
      let currentValue = null

      for (const line of lines) {
        if (line.length === 0) continue

        // Continuation line (starts with whitespace)
        if (/^\s/.test(line) && currentKey) {
          currentValue += ' ' + line.trim()
        } else {
          // Save previous header
          if (currentKey) {
            if (headers[currentKey]) {
              if (Array.isArray(headers[currentKey])) {
                headers[currentKey].push(currentValue)
              } else {
                headers[currentKey] = [headers[currentKey], currentValue]
              }
            } else {
              headers[currentKey] = currentValue
            }
          }

          const colonIndex = line.indexOf(':')
          if (colonIndex > 0) {
            currentKey = line.substring(0, colonIndex).trim()
            currentValue = line.substring(colonIndex + 1).trim()
          } else {
            currentKey = null
            currentValue = null
          }
        }
      }

      // Save last header
      if (currentKey) {
        if (headers[currentKey]) {
          if (Array.isArray(headers[currentKey])) {
            headers[currentKey].push(currentValue)
          } else {
            headers[currentKey] = [headers[currentKey], currentValue]
          }
        } else {
          headers[currentKey] = currentValue
        }
      }

      return headers
    },

    _updateUI() {
      if (typeof document === 'undefined') return

      const statusEl = document.getElementById('status')
      const userNameEl = document.getElementById('user-name')
      const folderNameEl = document.getElementById('folder-name')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')
      const eventsEl = document.getElementById('events')

      if (statusEl) {
        if (this._state.capturedEmail) {
          statusEl.innerHTML = '<span class="status-dot status-active"></span> Email selected'
        } else {
          statusEl.innerHTML = '<span class="status-dot status-idle"></span> No email selected'
        }
      }

      if (userNameEl) {
        userNameEl.textContent = this._state.userInfo.userName || 'Unknown user'
      }

      if (folderNameEl) {
        folderNameEl.textContent = this._state.userInfo.folderName || 'Unknown'
      }

      if (platformEl) {
        platformEl.textContent = this._state.userInfo.platform || 'Unknown'
      }

      if (versionEl) {
        versionEl.textContent = this._state.userInfo.version || 'Unknown'
      }

      if (eventsEl) {
        this._renderEvents(eventsEl)
      }
    },

    _renderEvents(container) {
      if (!this._state.events || this._state.events.length === 0) {
        container.innerHTML = '<div class="no-events">No events recorded yet.</div>'
        return
      }

      const reversed = [...this._state.events].reverse()
      container.innerHTML = reversed.map((evt) => {
        const detailStr = typeof evt.details === 'object'
          ? JSON.stringify(evt.details, null, 2)
          : String(evt.details || '')
        return `<div class="event-item">
          <div class="event-name">${this._escapeHtml(evt.name)}</div>
          <div class="event-timestamp">${evt.timestamp}</div>
          <div class="event-details">${this._escapeHtml(detailStr)}</div>
        </div>`
      }).join('')
    },

    _escapeHtml(str) {
      const div = document.createElement('div')
      div.textContent = str
      return div.innerHTML
    }
  }
}