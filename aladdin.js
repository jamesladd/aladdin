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
      items: {},
      userInfo: {
        name: '',
        platform: '',
        version: ''
      }
    },

    state() {
      return this._state
    },

    saveState() {
      try {
        localStorage.setItem('aladdin-state', JSON.stringify(this._state))
      } catch (e) {
        console.error('Failed to save state:', e)
      }
    },

    loadState() {
      try {
        const saved = localStorage.getItem('aladdin-state')
        if (saved) {
          const parsed = JSON.parse(saved)
          this._state = { ...this._state, ...parsed }
          this.updateUI()
        }
      } catch (e) {
        console.error('Failed to load state:', e)
      }
    },

    watchState() {
      if (typeof window === 'undefined') return

      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdin-state' && e.newValue) {
          try {
            const parsed = JSON.parse(e.newValue)
            this._state = { ...this._state, ...parsed }
            this.updateUI()
          } catch (err) {
            console.error('Failed to parse storage event:', err)
          }
        }
      })
    },

    event(name, details) {
      const timestamp = new Date().toISOString()
      const eventEntry = {
        name,
        details,
        timestamp
      }

      // Keep only last 10 events, most recent first
      this._state.events.unshift(eventEntry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      this.saveState()
      this.updateUI()
    },

    async initialize() {
      console.log('Aladdin initializing...')

      // Get user information
      await this.getUserInfo()

      // Add InboxAgent category
      await this.addInboxAgentCategory()

      // Register event handlers
      this.registerEventHandlers()

      // Get current item if one is selected
      if (this.Office.context.mailbox.item) {
        await this.captureCurrentItem()
      }

      this.updateUI()
      console.log('Aladdin initialized')
    },

    async getUserInfo() {
      try {
        const userProfile = this.Office.context.mailbox.userProfile
        this._state.userInfo.name = userProfile.displayName || userProfile.emailAddress || 'Unknown User'

        const platform = this.Office.context.platform
        const platformNames = {
          [this.Office.PlatformType.PC]: 'Windows Desktop',
          [this.Office.PlatformType.Mac]: 'Mac Desktop',
          [this.Office.PlatformType.OfficeOnline]: 'Office Online',
          [this.Office.PlatformType.iOS]: 'iOS',
          [this.Office.PlatformType.Android]: 'Android'
        }
        this._state.userInfo.platform = platformNames[platform] || 'Unknown Platform'

        const diagnostics = this.Office.context.diagnostics
        this._state.userInfo.version = diagnostics.version || 'Unknown Version'

        this.saveState()
        this.event('UserInfo', this._state.userInfo)
      } catch (e) {
        console.error('Failed to get user info:', e)
      }
    },

    async addInboxAgentCategory() {
      try {
        const mailbox = this.Office.context.mailbox
        if (mailbox.masterCategories && mailbox.masterCategories.addAsync) {
          mailbox.masterCategories.addAsync(
            [
              {
                displayName: 'InboxAgent',
                color: this.Office.MailboxEnums.CategoryColor.Preset9 // Closest to #009999
              }
            ],
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                this.event('CategoryAdded', { name: 'InboxAgent', color: '#009999' })
              } else {
                console.warn('Failed to add category:', result.error)
              }
            }
          )
        }
      } catch (e) {
        console.error('Failed to add InboxAgent category:', e)
      }
    },

    registerEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events
      if (mailbox.addHandlerAsync) {
        // ItemChanged - when user selects a different item
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          () => this.onItemChanged(),
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('ItemChanged handler registered')
            }
          }
        )

        // OfficeThemeChanged
        mailbox.addHandlerAsync(
          this.Office.EventType.OfficeThemeChanged,
          (args) => this.event('OfficeThemeChanged', args),
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('OfficeThemeChanged handler registered')
            }
          }
        )
      }

      // Item-level events - only if item exists
      this.registerItemEventHandlers()
    },

    registerItemEventHandlers() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.addHandlerAsync) return

      const itemEvents = [
        'RecipientsChanged',
        'AttachmentsChanged',
        'RecurrenceChanged',
        'AppointmentTimeChanged',
        'EnhancedLocationsChanged'
      ]

      itemEvents.forEach(eventType => {
        if (this.Office.EventType[eventType]) {
          item.addHandlerAsync(
            this.Office.EventType[eventType],
            (args) => {
              this.event(eventType, args)
              // Re-capture item on changes
              this.captureCurrentItem()
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log(`${eventType} handler registered`)
              }
            }
          )
        }
      })
    },

    async onItemChanged() {
      this.event('ItemChanged', { previousItemId: this._currentItemId })

      // Log the previous item before changing
      if (this._currentItemId && this._state.items[this._currentItemId]) {
        console.log('Previous item data:', JSON.stringify(this._state.items[this._currentItemId], null, 2))
      }

      // Register handlers for new item
      this.registerItemEventHandlers()

      // Capture new item
      await this.captureCurrentItem()
    },

    async captureCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) {
        this._currentItemId = null
        return
      }

      const itemId = item.itemId || `temp-${Date.now()}`

      // Log previous item if changing
      if (this._currentItemId && this._currentItemId !== itemId && this._state.items[this._currentItemId]) {
        console.log('Previous item data:', JSON.stringify(this._state.items[this._currentItemId], null, 2))
      }

      this._currentItemId = itemId

      const itemData = {
        itemId,
        itemType: item.itemType,
        subject: await this.getItemField('subject'),
        from: await this.getItemField('from'),
        to: await this.getItemField('to'),
        cc: await this.getItemField('cc'),
        categories: await this.getItemField('categories'),
        conversationId: item.conversationId,
        dateTimeCreated: item.dateTimeCreated?.toISOString(),
        dateTimeModified: item.dateTimeModified?.toISOString(),
        internetMessageId: item.internetMessageId,
        normalizedSubject: item.normalizedSubject,
        capturedAt: new Date().toISOString()
      }

      // Get flags
      if (item.getAsync) {
        itemData.flags = await this.getFlags(item)
      }

      // Get attachments
      itemData.attachments = this.getAttachmentInfo(item)

      // Get internet headers
      itemData.headers = await this.getInternetHeaders(item)

      this._state.items[itemId] = itemData
      this.saveState()
      this.event('ItemCaptured', { itemId, subject: itemData.subject })
    },

    async getItemField(fieldName) {
      const item = this.Office.context.mailbox.item
      if (!item) return null

      const field = item[fieldName]
      if (!field) return null

      // Check if getAsync is available for this field
      if (field.getAsync && typeof field.getAsync === 'function') {
        return new Promise((resolve) => {
          field.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(this.formatFieldValue(result.value))
            } else {
              resolve(null)
            }
          })
        })
      } else {
        // Use direct property access
        return this.formatFieldValue(field)
      }
    },

    formatFieldValue(value) {
      if (!value) return null

      // Handle EmailAddressDetails array
      if (Array.isArray(value)) {
        return value.map(v => {
          if (v.emailAddress) {
            return {
              name: v.displayName || '',
              email: v.emailAddress
            }
          }
          return v
        })
      }

      // Handle single EmailAddressDetails
      if (value.emailAddress) {
        return {
          name: value.displayName || '',
          email: value.emailAddress
        }
      }

      return value
    },

    async getFlags(item) {
      if (!item.getAsync) return null

      return new Promise((resolve) => {
        item.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const value = result.value
            resolve({
              flagStatus: value.flagStatus,
              startDate: value.startDate?.toISOString(),
              dueDate: value.dueDate?.toISOString(),
              completeDate: value.completeDate?.toISOString()
            })
          } else {
            resolve(null)
          }
        })
      })
    },

    getAttachmentInfo(item) {
      if (!item.attachments || item.attachments.length === 0) return []

      return item.attachments.map(att => ({
        id: att.id,
        name: att.name,
        attachmentType: att.attachmentType,
        size: att.size,
        contentType: att.contentType,
        isInline: att.isInline
      }))
    },

    async getInternetHeaders(item) {
      if (!item.getAllInternetHeadersAsync) return null

      return new Promise((resolve) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headersString = result.value
            const headersObj = this.parseInternetHeaders(headersString)
            resolve(headersObj)
          } else {
            resolve(null)
          }
        })
      })
    },

    parseInternetHeaders(headersString) {
      if (!headersString) return {}

      const headers = {}
      const lines = headersString.split('\n')
      let currentKey = null

      for (let line of lines) {
        // Check if line starts with whitespace (continuation of previous header)
        if (line.match(/^\s/) && currentKey) {
          headers[currentKey] += ' ' + line.trim()
        } else {
          // New header
          const colonIndex = line.indexOf(':')
          if (colonIndex > 0) {
            currentKey = line.substring(0, colonIndex).trim()
            const value = line.substring(colonIndex + 1).trim()
            headers[currentKey] = value
          }
        }
      }

      return headers
    },

    updateUI() {
      if (typeof document === 'undefined') return

      // Update user info
      const statusDiv = document.getElementById('status')
      if (statusDiv) {
        statusDiv.innerHTML = `
          <div class="user-info">
            <div class="user-name">${this._state.userInfo.name}</div>
            <div class="platform">Platform: ${this._state.userInfo.platform}</div>
            <div class="platform">Version: ${this._state.userInfo.version}</div>
          </div>
        `
      }

      // Update events list
      const eventsDiv = document.getElementById('events')
      if (eventsDiv) {
        if (this._state.events.length === 0) {
          eventsDiv.innerHTML = '<div class="no-events">No events yet</div>'
        } else {
          eventsDiv.innerHTML = this._state.events.map(event => `
            <div class="event-item">
              <div class="event-name">${this.escapeHtml(event.name)}</div>
              <div class="event-timestamp">${new Date(event.timestamp).toLocaleString()}</div>
              <div class="event-details">${this.escapeHtml(JSON.stringify(event.details, null, 2))}</div>
            </div>
          `).join('')
        }
      }
    },

    escapeHtml(text) {
      if (typeof text !== 'string') return text
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    }
  }
}