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
      userInfo: {
        userName: '',
        folder: '',
        platform: '',
        version: ''
      },
      capturedEmail: null
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
        const stored = localStorage.getItem('aladdin_state')
        if (stored) {
          const parsed = JSON.parse(stored)
          this._state = parsed
          if (this._state.capturedEmail && this._state.capturedEmail.itemId) {
            this._currentItemId = this._state.capturedEmail.itemId
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
          if (e.key === 'aladdin_state' && e.newValue) {
            try {
              this._state = JSON.parse(e.newValue)
              if (this._state.capturedEmail && this._state.capturedEmail.itemId) {
                this._currentItemId = this._state.capturedEmail.itemId
              } else {
                this._currentItemId = null
              }
              this.updateUI()
            } catch (err) {
              console.error('Failed to parse storage update', err)
            }
          }
        })
      }
    },

    _reloadAndCheckCapturedEmail() {
      try {
        const stored = localStorage.getItem('aladdin_state')
        if (stored) {
          const parsed = JSON.parse(stored)
          if (!parsed.capturedEmail) {
            this._state.capturedEmail = null
            this._currentItemId = null
            return null
          }
          return parsed.capturedEmail
        }
      } catch (e) {
        // fall through
      }
      return this._state.capturedEmail
    },

    event(name, details) {
      console.log('Event:', name)
      const entry = {
        name: name,
        details: details || {},
        timestamp: new Date().toISOString()
      }
      this._state.events.push(entry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(-10)
      }
      this.saveState()
      this.updateUI()
    },

    _isComposeMode() {
      const item = this.Office.context.mailbox.item
      return item && !item.itemId
    },

    _getCurrentItemId() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.itemId) return null
      try {
        return this.Office.context.mailbox.convertToRestId(
          item.itemId,
          this.Office.MailboxEnums.RestVersion.v2_0
        )
      } catch (e) {
        return item.itemId
      }
    },

    async _notifyPreviousIfNeeded() {
      const previousEmail = this._reloadAndCheckCapturedEmail()
      if (!previousEmail) return

      const mailbox = this.Office.context.mailbox
      const noItem = !mailbox.item
      const isCompose = this._isComposeMode()
      const currentId = this._getCurrentItemId()
      const previousId = previousEmail.itemId

      const isDifferentItem = currentId && previousId && currentId !== previousId

      if (noItem || isCompose || isDifferentItem) {
        this.event('deselection', {
          reason: noItem ? 'noItem' : isCompose ? 'compose' : 'differentItem',
          previousSubject: previousEmail.subject
        })
        await this.notify(previousEmail)
        this._state.capturedEmail = null
        this._currentItemId = null
        this.saveState()
      }
    },

    async initialize() {
      const mailbox = this.Office.context.mailbox
      const userProfile = mailbox.userProfile

      const userName = userProfile.displayName || userProfile.emailAddress || 'Unknown'
      const platform = this.Office.context.platform || 'Unknown'
      let version = 'Unknown'
      if (this.Office.context.diagnostics && this.Office.context.diagnostics.version) {
        version = this.Office.context.diagnostics.version
      }

      let folder = 'Inbox'
      if (mailbox.item && typeof mailbox.item.getAllInternetHeadersAsync === 'function') {
        try {
          const headers = await this._getAllInternetHeadersAsync(mailbox.item)
          if (headers) {
            const parsed = this._parseHeaders(headers)
            if (parsed['X-Folder']) {
              folder = parsed['X-Folder']
            } else if (parsed['X-Folder-Name']) {
              folder = parsed['X-Folder-Name']
            }
          }
        } catch (e) {
          // keep default
        }
      }

      this._state.userInfo = {
        userName: userName,
        folder: folder,
        platform: String(platform),
        version: version
      }
      this.saveState()
      this.updateUI()

      this.event('initialized', { userName, folder, platform: String(platform), version })

      await this._notifyPreviousIfNeeded()

      this._registerMailboxHandlers()

      if (mailbox.item && !this._isComposeMode()) {
        this._registerItemHandlers()
        await this._captureCurrentItem()
      }

      await this._initCategories()

      this._updateStatus('Ready')
    },

    async _initCategories() {
      try {
        const categories = await this.getCategories(this._state.userInfo)
        if (categories && categories.length > 0 && this.Office.context.mailbox.masterCategories) {
          this.Office.context.mailbox.masterCategories.addAsync(categories, (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              this.event('categoriesAdded', { count: categories.length })
            } else {
              this.event('categoriesAddFailed', { error: result.error ? result.error.message : 'unknown' })
            }
          })
        }
      } catch (e) {
        this.event('categoriesError', { error: e.message })
      }
    },

    _registerMailboxHandlers() {
      const mailbox = this.Office.context.mailbox
      const EventType = this.Office.EventType

      if (EventType.ItemChanged) {
        mailbox.addHandlerAsync(EventType.ItemChanged, (eventArgs) => {
          this._onItemChanged(eventArgs)
        })
      }

      if (EventType.OfficeThemeChanged) {
        mailbox.addHandlerAsync(EventType.OfficeThemeChanged, (eventArgs) => {
          this.event('OfficeThemeChanged', eventArgs || {})
        })
      }
    },

    _registerItemHandlers() {
      const item = this.Office.context.mailbox.item
      if (!item) return
      if (this._itemHandlersRegistered) return
      this._itemHandlersRegistered = true

      const EventType = this.Office.EventType
      const itemEvents = [
        'RecipientsChanged',
        'AttachmentsChanged',
        'RecurrenceChanged',
        'AppointmentTimeChanged',
        'EnhancedLocationsChanged'
      ]

      itemEvents.forEach((eventName) => {
        if (EventType[eventName]) {
          item.addHandlerAsync(EventType[eventName], (eventArgs) => {
            this.event(eventName, eventArgs || {})
            if (eventName === 'RecipientsChanged' || eventName === 'AttachmentsChanged') {
              this._captureCurrentItem()
            }
          })
        }
      })
    },

    async _onItemChanged(eventArgs) {
      this.event('ItemChanged', eventArgs || {})

      const previousEmail = this._reloadAndCheckCapturedEmail()

      if (previousEmail) {
        await this.notify(previousEmail)
        this._state.capturedEmail = null
        this._currentItemId = null
        this.saveState()
      }

      this._itemHandlersRegistered = false
      const mailbox = this.Office.context.mailbox

      if (mailbox.item && !this._isComposeMode()) {
        this._registerItemHandlers()
        await this._captureCurrentItem()
      } else {
        this._currentItemId = null
        this._state.capturedEmail = null
        this.saveState()
        this.updateUI()
      }
    },

    async _captureCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const emailData = {}

      emailData.to = await this._getRecipientField(item, 'to')
      emailData.from = await this._getFromField(item)
      emailData.cc = await this._getRecipientField(item, 'cc')
      emailData.subject = await this._getStringField(item, 'subject')
      emailData.attachments = this._getAttachmentNames(item)
      emailData.categories = await this._getCategoriesField(item)
      emailData.importance = item.importance || 'normal'
      emailData.sentiment = this._deriveSentiment(emailData.importance)

      emailData.itemId = null
      if (item.itemId) {
        try {
          emailData.itemId = this.Office.context.mailbox.convertToRestId(
            item.itemId,
            this.Office.MailboxEnums.RestVersion.v2_0
          )
        } catch (e) {
          emailData.itemId = item.itemId
        }
      }

      emailData.internetHeaders = {}
      if (typeof item.getAllInternetHeadersAsync === 'function') {
        try {
          const headersStr = await this._getAllInternetHeadersAsync(item)
          if (headersStr) {
            emailData.internetHeaders = this._parseHeaders(headersStr)
          }
        } catch (e) {
          // ignore
        }
      }

      this._currentItemId = emailData.itemId || item.itemId || ('item_' + Date.now())
      this._state.capturedEmail = emailData
      this.saveState()
      this.updateUI()

      this.event('emailCaptured', { subject: emailData.subject, itemId: this._currentItemId })
    },

    async _getRecipientField(item, fieldName) {
      const field = item[fieldName]
      if (!field) return []

      if (typeof field.getAsync === 'function') {
        return new Promise((resolve) => {
          field.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(this._formatRecipients(result.value))
            } else {
              resolve([])
            }
          })
        })
      }

      if (Array.isArray(field)) {
        return this._formatRecipients(field)
      }

      return []
    },

    async _getFromField(item) {
      const from = item.from
      if (!from) return null

      if (typeof from.getAsync === 'function') {
        return new Promise((resolve) => {
          from.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              const val = result.value
              resolve({ displayName: val.displayName, emailAddress: val.emailAddress })
            } else {
              resolve(null)
            }
          })
        })
      }

      if (from.displayName || from.emailAddress) {
        return { displayName: from.displayName, emailAddress: from.emailAddress }
      }

      return null
    },

    async _getStringField(item, fieldName) {
      const field = item[fieldName]
      if (field === undefined || field === null) return ''

      if (typeof field === 'object' && typeof field.getAsync === 'function') {
        return new Promise((resolve) => {
          field.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || '')
            } else {
              resolve('')
            }
          })
        })
      }

      return String(field)
    },

    async _getCategoriesField(item) {
      if (!item.categories) return []

      if (typeof item.categories.getAsync === 'function') {
        return new Promise((resolve) => {
          item.categories.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || [])
            } else {
              resolve([])
            }
          })
        })
      }

      if (Array.isArray(item.categories)) {
        return item.categories
      }

      return []
    },

    _formatRecipients(recipients) {
      if (!recipients || !Array.isArray(recipients)) return []
      return recipients.map((r) => ({
        displayName: r.displayName || '',
        emailAddress: r.emailAddress || ''
      }))
    },

    _getAttachmentNames(item) {
      if (!item.attachments || !Array.isArray(item.attachments)) return []
      return item.attachments.map((a) => a.name || 'unnamed')
    },

    _deriveSentiment(importance) {
      if (importance === 'high') return 'urgent'
      if (importance === 'low') return 'low'
      return 'normal'
    },

    _getAllInternetHeadersAsync(item) {
      return new Promise((resolve, reject) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value)
          } else {
            reject(new Error(result.error ? result.error.message : 'Failed to get headers'))
          }
        })
      })
    },

    _parseHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers

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

    async notify(emailData) {
      if (!emailData) return

      this.event('notify', { subject: emailData.subject, itemId: emailData.itemId })

      try {
        const response = await fetch('/api/inboxnotify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email: emailData,
            userInfo: this._state.userInfo,
            timestamp: new Date().toISOString()
          })
        })
        if (!response.ok) {
          this.event('notifyFailed', { status: response.status })
        } else {
          this.event('notifySuccess', { itemId: emailData.itemId })
        }
      } catch (e) {
        this.event('notifyError', { error: e.message })
      }
    },

    async getCategories(userInfo) {
      const defaultCategories = [
        { displayName: 'InboxAgent', color: 'Preset0' }
      ]

      try {
        const response = await fetch('/api/inboxinit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            userInfo: userInfo || this._state.userInfo,
            timestamp: new Date().toISOString()
          })
        })
        if (response.ok) {
          const data = await response.json()
          if (data && Array.isArray(data.categories) && data.categories.length > 0) {
            return data.categories
          }
        }
      } catch (e) {
        // fall through to default
      }

      return defaultCategories
    },

    _updateStatus(text) {
      if (typeof document === 'undefined') return
      const el = document.getElementById('status')
      if (el) el.textContent = text
    },

    updateUI() {
      if (typeof document === 'undefined') return

      const userNameEl = document.getElementById('user-name')
      const folderEl = document.getElementById('user-folder')
      const platformEl = document.getElementById('user-platform')
      const versionEl = document.getElementById('user-version')

      if (userNameEl) userNameEl.textContent = this._state.userInfo.userName || ''
      if (folderEl) folderEl.textContent = 'Folder: ' + (this._state.userInfo.folder || '')
      if (platformEl) platformEl.textContent = 'Platform: ' + (this._state.userInfo.platform || '')
      if (versionEl) versionEl.textContent = 'Version: ' + (this._state.userInfo.version || '')

      const eventsEl = document.getElementById('events')
      if (!eventsEl) return

      if (this._state.events.length === 0) {
        eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      const eventsHtml = this._state.events
        .slice()
        .reverse()
        .map((evt) => {
          const detailsStr = typeof evt.details === 'object'
            ? JSON.stringify(evt.details, null, 2)
            : String(evt.details)
          return '<div class="event-item">' +
            '<div class="event-name">' + this._escapeHtml(evt.name) + '</div>' +
            '<div class="event-timestamp">' + this._escapeHtml(evt.timestamp) + '</div>' +
            '<div class="event-details">' + this._escapeHtml(detailsStr) + '</div>' +
            '</div>'
        })
        .join('')

      eventsEl.innerHTML = eventsHtml
    },

    _escapeHtml(str) {
      if (!str) return ''
      return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
    }
  }
}