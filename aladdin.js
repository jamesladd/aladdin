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
      userName: '',
      platform: '',
      officeVersion: '',
    },

    state() {
      return this._state
    },

    saveState() {
      try {
        const stateJson = JSON.stringify(this._state)
        localStorage.setItem('aladdin_state', stateJson)
        console.log('State saved:', this._state)
      } catch (error) {
        console.error('Error saving state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdin_state')
        if (stateJson) {
          const loadedState = JSON.parse(stateJson)
          this._state = { ...this._state, ...loadedState }
          console.log('State loaded:', this._state)
          this.updateUI()
        }
      } catch (error) {
        console.error('Error loading state:', error)
      }
    },

    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === 'aladdin_state' && e.newValue) {
            try {
              const newState = JSON.parse(e.newValue)
              this._state = { ...this._state, ...newState }
              console.log('State updated from storage:', this._state)
              this.updateUI()
            } catch (error) {
              console.error('Error watching state:', error)
            }
          }
        })
      }
    },

    event(name, details) {
      const timestamp = new Date().toISOString()
      const eventData = {
        name,
        details,
        timestamp
      }

      // Keep only last 10 events, most recent first
      this._state.events.unshift(eventData)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      console.log('Event recorded:', eventData)
      this.saveState()
      this.updateUI()
    },

    updateUI() {
      this.renderUserInfo()
      this.renderEvents()
    },

    renderUserInfo() {
      const userInfoDiv = document.getElementById('user-info')
      if (userInfoDiv) {
        userInfoDiv.innerHTML = `
          <div class="user-name">${this._state.userName || 'Loading...'}</div>
          <div class="platform">Platform: ${this._state.platform || 'Loading...'}</div>
          <div class="office-version">Office Version: ${this._state.officeVersion || 'Loading...'}</div>
        `
      }
    },

    renderEvents() {
      const eventsDiv = document.getElementById('events')
      if (!eventsDiv) return

      if (this._state.events.length === 0) {
        eventsDiv.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      const eventsHtml = this._state.events.map(event => `
        <div class="event-item">
          <div class="event-name">${this.escapeHtml(event.name)}</div>
          <div class="event-timestamp">${new Date(event.timestamp).toLocaleString()}</div>
          <div class="event-details">${this.escapeHtml(JSON.stringify(event.details, null, 2))}</div>
        </div>
      `).join('')

      eventsDiv.innerHTML = eventsHtml
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    async initialize() {
      console.log('Initializing Aladdin...')

      // Get user context information
      await this.loadUserContext()

      // Register all event handlers
      this.registerEventHandlers()

      // Load current item if available
      if (this.Office.context.mailbox.item) {
        await this.handleItemChanged()
      }

      this.updateUI()
    },

    async loadUserContext() {
      try {
        // Get user name
        const userProfile = this.Office.context.mailbox.userProfile
        this._state.userName = userProfile.displayName || userProfile.emailAddress || 'Unknown User'

        // Get platform
        this._state.platform = this.Office.context.platform || 'Unknown Platform'

        // Get Office version
        this._state.officeVersion = this.Office.context.diagnostics?.version || 'Unknown Version'

        this.event('UserContextLoaded', {
          userName: this._state.userName,
          platform: this._state.platform,
          officeVersion: this._state.officeVersion
        })

        this.saveState()
      } catch (error) {
        console.error('Error loading user context:', error)
      }
    },

    registerEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events - register on mailbox
      if (mailbox.addHandlerAsync) {
        // ItemChanged - when user selects different item
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          () => this.handleItemChanged(),
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('ItemChanged handler registered')
            }
          }
        )

        // OfficeThemeChanged - when Office theme changes
        if (this.Office.EventType.OfficeThemeChanged) {
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
      }

      // Item-level events - register on item if available
      this.registerItemLevelEvents()
    },

    registerItemLevelEvents() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.addHandlerAsync) return

      // RecipientsChanged - when To, CC, or BCC recipients change
      if (this.Office.EventType.RecipientsChanged) {
        item.addHandlerAsync(
          this.Office.EventType.RecipientsChanged,
          (args) => {
            this.event('RecipientsChanged', {
              type: args.changedRecipientFields,
              itemId: this._currentItemId
            })
            this.captureItemFields()
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('RecipientsChanged handler registered')
            }
          }
        )
      }

      // AttachmentsChanged - when attachments are added or removed
      if (this.Office.EventType.AttachmentsChanged) {
        item.addHandlerAsync(
          this.Office.EventType.AttachmentsChanged,
          (args) => {
            this.event('AttachmentsChanged', {
              attachmentStatus: args.attachmentStatus,
              attachmentDetails: args.attachmentDetails,
              itemId: this._currentItemId
            })
            this.captureItemFields()
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('AttachmentsChanged handler registered')
            }
          }
        )
      }

      // RecurrenceChanged - when recurrence pattern changes (for appointments)
      if (this.Office.EventType.RecurrenceChanged) {
        item.addHandlerAsync(
          this.Office.EventType.RecurrenceChanged,
          (args) => {
            this.event('RecurrenceChanged', {
              recurrence: args.recurrence,
              itemId: this._currentItemId
            })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('RecurrenceChanged handler registered')
            }
          }
        )
      }

      // AppointmentTimeChanged - when appointment time changes
      if (this.Office.EventType.AppointmentTimeChanged) {
        item.addHandlerAsync(
          this.Office.EventType.AppointmentTimeChanged,
          (args) => {
            this.event('AppointmentTimeChanged', {
              start: args.start,
              end: args.end,
              itemId: this._currentItemId
            })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('AppointmentTimeChanged handler registered')
            }
          }
        )
      }

      // EnhancedLocationsChanged - when enhanced locations change
      if (this.Office.EventType.EnhancedLocationsChanged) {
        item.addHandlerAsync(
          this.Office.EventType.EnhancedLocationsChanged,
          (args) => {
            this.event('EnhancedLocationsChanged', {
              enhancedLocations: args.enhancedLocations,
              itemId: this._currentItemId
            })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('EnhancedLocationsChanged handler registered')
            }
          }
        )
      }
    },

    async handleItemChanged() {
      const item = this.Office.context.mailbox.item

      // Check if item was deselected
      if (!item) {
        if (this._currentItemId) {
          this.event('ItemDeselected', {
            previousItemId: this._currentItemId
          })
          this._currentItemId = null
        }
        return
      }

      // Check if this is a new item selection
      const newItemId = item.itemId
      const isNewSelection = this._currentItemId !== newItemId

      if (isNewSelection) {
        const previousItemId = this._currentItemId
        this._currentItemId = newItemId

        this.event('ItemSelected', {
          itemId: this._currentItemId,
          previousItemId: previousItemId,
          itemType: item.itemType
        })

        // Register item-level events for the new item
        this.registerItemLevelEvents()

        // Capture all fields of the selected item
        await this.captureItemFields()
      }
    },

    async captureItemFields() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      try {
        const fields = {
          itemId: this._currentItemId,
          itemType: item.itemType
        }

        // Use async methods if available, otherwise use sync
        if (item.subject && item.subject.getAsync) {
          // Async API available
          await Promise.all([
            this.getAsyncValue(item.subject, 'subject', fields),
            this.getAsyncValue(item.to, 'to', fields),
            this.getAsyncValue(item.from, 'from', fields),
            this.getAsyncValue(item.cc, 'cc', fields),
            this.getAsyncValue(item.categories, 'categories', fields),
          ])
        } else {
          // Fallback to sync API
          fields.subject = item.subject || ''
          fields.to = item.to || []
          fields.from = item.from || {}
          fields.cc = item.cc || []
          fields.categories = item.categories || []
        }

        // Get flags (sync property)
        if (item.itemType === this.Office.MailboxEnums.ItemType.Message) {
          fields.flags = {
            importance: item.importance,
            sensitivity: item.sensitivity
          }
        }

        // Get attachments
        fields.attachments = (item.attachments || []).map(att => ({
          id: att.id,
          name: att.name,
          attachmentType: att.attachmentType,
          size: att.size,
          contentType: att.contentType,
          isInline: att.isInline
        }))

        // Get internet headers
        if (item.getAllInternetHeadersAsync) {
          await new Promise((resolve) => {
            item.getAllInternetHeadersAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                fields.headers = this.parseInternetHeaders(result.value)
              }
              resolve()
            })
          })
        }

        this.event('ItemFieldsCaptured', fields)
      } catch (error) {
        console.error('Error capturing item fields:', error)
        this.event('ItemFieldsCaptureError', { error: error.message })
      }
    },

    async getAsyncValue(property, fieldName, targetObject) {
      if (!property || !property.getAsync) return

      return new Promise((resolve) => {
        property.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            targetObject[fieldName] = result.value
          }
          resolve()
        })
      })
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
        } else if (line.includes(':')) {
          // Save previous header if exists
          if (currentKey) {
            headers[currentKey] = currentValue.trim()
          }
          // Start new header
          const colonIndex = line.indexOf(':')
          currentKey = line.substring(0, colonIndex).trim()
          currentValue = line.substring(colonIndex + 1).trim()
        }
      }

      // Save last header
      if (currentKey) {
        headers[currentKey] = currentValue.trim()
      }

      return headers
    }
  }
}