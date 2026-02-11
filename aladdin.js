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
        const stateJson = JSON.stringify(this._state)
        localStorage.setItem('aladdin-state', stateJson)
        console.log('State saved:', this._state)
      } catch (error) {
        console.error('Error saving state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdin-state')
        if (stateJson) {
          this._state = JSON.parse(stateJson)
          console.log('State loaded:', this._state)
        }
      } catch (error) {
        console.error('Error loading state:', error)
      }
    },

    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === 'aladdin-state' && e.newValue) {
            try {
              this._state = JSON.parse(e.newValue)
              console.log('State updated from storage:', this._state)
              this.updateUI()
            } catch (error) {
              console.error('Error parsing storage event:', error)
            }
          }
        })
      }
    },

    event(name, details) {
      const event = {
        name,
        details,
        timestamp: new Date().toISOString()
      }

      // Add to beginning of events array
      this._state.events.unshift(event)

      // Keep only last 10 events
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      console.log('Event recorded:', event)
      this.saveState()
      this.updateUI()
    },

    updateUI() {
      this.renderUserInfo()
      this.renderEvents()
    },

    renderUserInfo() {
      const userInfoEl = document.getElementById('user-info')
      if (!userInfoEl) return

      const { name, platform, version } = this._state.userInfo
      userInfoEl.innerHTML = `
        <div class="user-name">${name || 'Loading...'}</div>
        <div class="platform">Platform: ${platform || 'Unknown'}</div>
        <div class="platform">Version: ${version || 'Unknown'}</div>
      `
    },

    renderEvents() {
      const eventsEl = document.getElementById('events')
      if (!eventsEl) return

      if (this._state.events.length === 0) {
        eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      eventsEl.innerHTML = this._state.events.map(event => `
        <div class="event-item">
          <div class="event-name">${this.escapeHtml(event.name)}</div>
          <div class="event-timestamp">${new Date(event.timestamp).toLocaleString()}</div>
          <div class="event-details">${this.escapeHtml(JSON.stringify(event.details, null, 2))}</div>
        </div>
      `).join('')
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    async getUserInfo() {
      try {
        const userProfile = this.Office.context.mailbox.userProfile
        this._state.userInfo.name = userProfile.displayName || userProfile.emailAddress || 'Unknown User'

        const platform = this.Office.context.platform || this.Office.context.host || 'Unknown'
        this._state.userInfo.platform = this.getPlatformName(platform)

        const diagnostics = this.Office.context.diagnostics
        this._state.userInfo.version = diagnostics ? diagnostics.version : 'Unknown'

        this.saveState()
        this.updateUI()

        this.event('UserInfoLoaded', {
          name: this._state.userInfo.name,
          platform: this._state.userInfo.platform,
          version: this._state.userInfo.version
        })
      } catch (error) {
        console.error('Error getting user info:', error)
        this.event('UserInfoError', { error: error.message })
      }
    },

    getPlatformName(platform) {
      const platformMap = {
        0: 'PC',
        1: 'OfficeOnline',
        2: 'Mac',
        3: 'iOS',
        4: 'Android',
        5: 'Universal'
      }
      return platformMap[platform] || String(platform)
    },

    parseInternetHeaders(headersString) {
      const headers = {}
      if (!headersString) return headers

      const lines = headersString.split('\r\n')
      let currentKey = null
      let currentValue = ''

      for (const line of lines) {
        if (line.match(/^\s/) && currentKey) {
          // Continuation of previous header
          currentValue += ' ' + line.trim()
        } else {
          // Save previous header
          if (currentKey) {
            headers[currentKey] = currentValue.trim()
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
        headers[currentKey] = currentValue.trim()
      }

      return headers
    },

    async captureItemFields(item) {
      if (!item) return null

      const itemData = {
        itemId: item.itemId,
        conversationId: item.conversationId,
        subject: null,
        from: null,
        to: null,
        cc: null,
        categories: null,
        flags: null,
        attachments: [],
        internetHeaders: {}
      }

      try {
        // Check if getAsync is available for each field
        if (item.subject && typeof item.subject.getAsync === 'function') {
          itemData.subject = await this.getAsync(item.subject)
        } else {
          itemData.subject = item.subject
        }

        if (item.from && typeof item.from.getAsync === 'function') {
          itemData.from = await this.getAsync(item.from)
        } else {
          itemData.from = item.from
        }

        if (item.to && typeof item.to.getAsync === 'function') {
          itemData.to = await this.getAsync(item.to)
        } else {
          itemData.to = item.to
        }

        if (item.cc && typeof item.cc.getAsync === 'function') {
          itemData.cc = await this.getAsync(item.cc)
        } else {
          itemData.cc = item.cc
        }

        if (item.categories && typeof item.categories.getAsync === 'function') {
          itemData.categories = await this.getAsync(item.categories)
        } else {
          itemData.categories = item.categories
        }

        // Get attachments
        if (item.attachments) {
          itemData.attachments = item.attachments.map(att => ({
            id: att.id,
            name: att.name,
            size: att.size,
            attachmentType: att.attachmentType,
            isInline: att.isInline,
            contentType: att.contentType
          }))
        }

        // Get internet headers
        if (item.getAllInternetHeadersAsync) {
          const headersString = await this.getAllInternetHeadersAsync(item)
          itemData.internetHeaders = this.parseInternetHeaders(headersString)
        }

      } catch (error) {
        console.error('Error capturing item fields:', error)
        this.event('ItemFieldsError', { error: error.message })
      }

      return itemData
    },

    getAsync(property) {
      return new Promise((resolve, reject) => {
        property.getAsync((asyncResult) => {
          if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(asyncResult.value)
          } else {
            reject(new Error(asyncResult.error.message))
          }
        })
      })
    },

    getAllInternetHeadersAsync(item) {
      return new Promise((resolve, reject) => {
        item.getAllInternetHeadersAsync((asyncResult) => {
          if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(asyncResult.value)
          } else {
            reject(new Error(asyncResult.error.message))
          }
        })
      })
    },

    async handleItemChanged(eventArgs) {
      this.event('ItemChanged', { eventType: eventArgs.type })

      const item = this.Office.context.mailbox.item
      if (!item) {
        this._currentItemId = null
        return
      }

      const newItemId = item.itemId

      // Log previous item before changing
      if (this._currentItemId && this._currentItemId !== newItemId) {
        const previousItem = this._state.items[this._currentItemId]
        if (previousItem) {
          console.log('Previous item before change:', JSON.stringify(previousItem))
        }
      }

      // Update current item ID
      this._currentItemId = newItemId

      // Capture new item fields
      const itemData = await this.captureItemFields(item)
      if (itemData) {
        this._state.items[newItemId] = itemData
        this.saveState()
        this.event('ItemCaptured', { itemId: newItemId })
      }
    },

    registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events
      mailbox.addHandlerAsync(
        this.Office.EventType.ItemChanged,
        this.handleItemChanged.bind(this),
        (asyncResult) => {
          if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('ItemChanged handler registered')
            this.event('HandlerRegistered', { eventType: 'ItemChanged' })
          } else {
            console.error('Failed to register ItemChanged handler:', asyncResult.error)
          }
        }
      )

      mailbox.addHandlerAsync(
        this.Office.EventType.OfficeThemeChanged,
        (eventArgs) => {
          this.event('OfficeThemeChanged', { eventType: eventArgs.type })
        },
        (asyncResult) => {
          if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('OfficeThemeChanged handler registered')
            this.event('HandlerRegistered', { eventType: 'OfficeThemeChanged' })
          }
        }
      )
    },

    registerItemEvents() {
      const item = this.Office.context.mailbox.item
      if (!item) {
        console.log('No item available for event registration')
        return
      }

      // Item-level events
      const itemEvents = [
        { type: this.Office.EventType.RecipientsChanged, name: 'RecipientsChanged' },
        { type: this.Office.EventType.AttachmentsChanged, name: 'AttachmentsChanged' },
        { type: this.Office.EventType.RecurrenceChanged, name: 'RecurrenceChanged' },
        { type: this.Office.EventType.AppointmentTimeChanged, name: 'AppointmentTimeChanged' },
        { type: this.Office.EventType.EnhancedLocationsChanged, name: 'EnhancedLocationsChanged' }
      ]

      itemEvents.forEach(({ type, name }) => {
        if (type) {
          item.addHandlerAsync(
            type,
            (eventArgs) => {
              this.event(name, { eventType: eventArgs.type })
            },
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log(`${name} handler registered`)
                this.event('HandlerRegistered', { eventType: name })
              }
            }
          )
        }
      })
    },

    async addInboxAgentCategory() {
      try {
        const mailbox = this.Office.context.mailbox
        if (!mailbox.masterCategories) {
          console.log('Master categories not supported')
          return
        }

        const categories = await new Promise((resolve, reject) => {
          mailbox.masterCategories.getAsync((asyncResult) => {
            if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(asyncResult.value)
            } else {
              reject(new Error(asyncResult.error.message))
            }
          })
        })

        // Check if InboxAgent category already exists
        const existingCategory = categories.find(cat => cat.displayName === 'InboxAgent')
        if (existingCategory) {
          console.log('InboxAgent category already exists')
          this.event('CategoryExists', { displayName: 'InboxAgent' })
          return
        }

        // Add new category
        await new Promise((resolve, reject) => {
          mailbox.masterCategories.addAsync(
            [{ displayName: 'InboxAgent', color: this.Office.MailboxEnums.CategoryColor.Preset9 }],
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve()
              } else {
                reject(new Error(asyncResult.error.message))
              }
            }
          )
        })

        console.log('InboxAgent category added')
        this.event('CategoryAdded', { displayName: 'InboxAgent', color: '#f54900' })

      } catch (error) {
        console.error('Error adding InboxAgent category:', error)
        this.event('CategoryAddError', { error: error.message })
      }
    },

    async initialize() {
      console.log('Initializing Aladdin add-in')
      this.event('Initialize', { timestamp: new Date().toISOString() })

      // Get user info
      await this.getUserInfo()

      // Register mailbox-level events
      this.registerMailboxEvents()

      // Register item-level events if item exists
      if (this.Office.context.mailbox.item) {
        this.registerItemEvents()

        // Capture current item
        await this.handleItemChanged({ type: 'Initial' })
      }

      // Add InboxAgent category
      await this.addInboxAgentCategory()

      this.updateUI()
      console.log('Aladdin add-in initialized')
    }
  }
}