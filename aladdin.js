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
    _state: {
      events: [],
      userInfo: null,
      platform: null,
      version: null,
    },
    state() {
      return this._state
    },
    saveState() {
      try {
        localStorage.setItem('aladdin-state', JSON.stringify(this._state))
      } catch (error) {
        console.error('Failed to save state:', error)
      }
    },
    loadState() {
      try {
        const stored = localStorage.getItem('aladdin-state')
        if (stored) {
          const parsed = JSON.parse(stored)
          this._state = { ...this._state, ...parsed }
          this.updateUI()
        }
      } catch (error) {
        console.error('Failed to load state:', error)
      }
    },
    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === 'aladdin-state' && e.newValue) {
            try {
              const parsed = JSON.parse(e.newValue)
              this._state = { ...this._state, ...parsed }
              this.updateUI()
            } catch (error) {
              console.error('Failed to parse storage event:', error)
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

      // Add to front of array
      this._state.events.unshift(event)

      // Keep only last 10 events
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      this.saveState()
      this.updateUI()
    },
    updateUI() {
      this.renderUserInfo()
      this.renderEvents()
    },
    renderUserInfo() {
      const userInfoDiv = document.getElementById('user-info')
      if (!userInfoDiv) return

      const { userInfo, platform, version } = this._state

      if (!userInfo && !platform && !version) {
        userInfoDiv.innerHTML = '<div class="user-info"><div class="user-name">Loading user information...</div></div>'
        return
      }

      userInfoDiv.innerHTML = `
        <div class="user-info">
          <div class="user-name">${userInfo || 'Unknown User'}</div>
          <div class="platform">Platform: ${platform || 'Unknown'}</div>
          <div class="version">Version: ${version || 'Unknown'}</div>
        </div>
      `
    },
    renderEvents() {
      const eventsDiv = document.getElementById('events')
      if (!eventsDiv) return

      const { events } = this._state

      if (events.length === 0) {
        eventsDiv.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      eventsDiv.innerHTML = events.map(event => `
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
    async initialize() {
      console.log('Initializing Aladdin...')

      // Get user information
      await this.getUserInfo()

      // Get platform and version
      this.getPlatformInfo()

      // Register event handlers
      this.registerEventHandlers()

      // Capture current email item if available
      if (this.Office.context.mailbox.item) {
        await this.captureEmailItem()
      }

      // Initial UI render
      this.updateUI()

      console.log('Aladdin initialized')
    },
    async getUserInfo() {
      try {
        const userProfile = this.Office.context.mailbox.userProfile
        this._state.userInfo = userProfile.displayName || userProfile.emailAddress || 'Unknown User'
        this.saveState()
      } catch (error) {
        console.error('Failed to get user info:', error)
        this._state.userInfo = 'Unknown User'
      }
    },
    getPlatformInfo() {
      try {
        const diagnostics = this.Office.context.diagnostics
        this._state.platform = diagnostics.platform || 'Unknown'
        this._state.version = diagnostics.version || 'Unknown'
        this.saveState()
      } catch (error) {
        console.error('Failed to get platform info:', error)
        this._state.platform = 'Unknown'
        this._state.version = 'Unknown'
      }
    },
    registerEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events
      if (mailbox.addHandlerAsync) {
        // ItemChanged - when user selects a different item
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          async () => {
            this.event('ItemChanged', { message: 'User selected a different item' })
            if (mailbox.item) {
              await this.captureEmailItem()
            }
          }
        )

        // OfficeThemeChanged
        mailbox.addHandlerAsync(
          this.Office.EventType.OfficeThemeChanged,
          (args) => {
            this.event('OfficeThemeChanged', args)
          }
        )
      }

      // Item-level events - only register if item exists
      if (mailbox.item && mailbox.item.addHandlerAsync) {
        // RecipientsChanged
        if (this.Office.EventType.RecipientsChanged) {
          mailbox.item.addHandlerAsync(
            this.Office.EventType.RecipientsChanged,
            async (args) => {
              this.event('RecipientsChanged', args)
              await this.captureEmailItem()
            }
          )
        }

        // AttachmentsChanged
        if (this.Office.EventType.AttachmentsChanged) {
          mailbox.item.addHandlerAsync(
            this.Office.EventType.AttachmentsChanged,
            async (args) => {
              this.event('AttachmentsChanged', args)
              await this.captureEmailItem()
            }
          )
        }

        // RecurrenceChanged (for appointments)
        if (this.Office.EventType.RecurrenceChanged) {
          mailbox.item.addHandlerAsync(
            this.Office.EventType.RecurrenceChanged,
            (args) => {
              this.event('RecurrenceChanged', args)
            }
          )
        }

        // AppointmentTimeChanged (for appointments)
        if (this.Office.EventType.AppointmentTimeChanged) {
          mailbox.item.addHandlerAsync(
            this.Office.EventType.AppointmentTimeChanged,
            (args) => {
              this.event('AppointmentTimeChanged', args)
            }
          )
        }

        // EnhancedLocationsChanged (for appointments)
        if (this.Office.EventType.EnhancedLocationsChanged) {
          mailbox.item.addHandlerAsync(
            this.Office.EventType.EnhancedLocationsChanged,
            (args) => {
              this.event('EnhancedLocationsChanged', args)
            }
          )
        }
      }
    },
    async captureEmailItem() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      try {
        const itemData = {
          itemType: item.itemType,
          subject: item.subject,
          dateTimeCreated: item.dateTimeCreated?.toISOString(),
          dateTimeModified: item.dateTimeModified?.toISOString(),
        }

        // Get recipients using async if available
        if (item.to) {
          itemData.to = await this.getRecipients(item.to)
        }

        if (item.from && item.from.getAsync) {
          itemData.from = await this.getFrom(item.from)
        }

        if (item.cc) {
          itemData.cc = await this.getRecipients(item.cc)
        }

        // Get categories
        if (item.categories) {
          itemData.categories = await this.getCategories(item.categories)
        }

        // Get attachments
        if (item.attachments) {
          itemData.attachments = item.attachments.map(att => ({
            id: att.id,
            name: att.name,
            size: att.size,
            attachmentType: att.attachmentType,
            isInline: att.isInline
          }))
        }

        // Get internet headers
        if (item.getAllInternetHeadersAsync) {
          itemData.internetHeaders = await this.getInternetHeaders()
        }

        this.event('EmailItemCaptured', itemData)
      } catch (error) {
        console.error('Failed to capture email item:', error)
        this.event('EmailItemCaptureError', { error: error.message })
      }
    },
    async getRecipients(recipientProperty) {
      return new Promise((resolve) => {
        // Check if getAsync is available
        if (recipientProperty.getAsync) {
          recipientProperty.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value.map(r => ({
                displayName: r.displayName,
                emailAddress: r.emailAddress
              })))
            } else {
              console.error('Failed to get recipients:', result.error)
              resolve([])
            }
          })
        } else {
          // Fallback for synchronous access
          resolve([])
        }
      })
    },
    async getFrom(fromProperty) {
      return new Promise((resolve) => {
        if (fromProperty.getAsync) {
          fromProperty.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve({
                displayName: result.value.displayName,
                emailAddress: result.value.emailAddress
              })
            } else {
              console.error('Failed to get from:', result.error)
              resolve(null)
            }
          })
        } else {
          resolve(null)
        }
      })
    },
    async getCategories(categoriesProperty) {
      return new Promise((resolve) => {
        if (categoriesProperty.getAsync) {
          categoriesProperty.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value)
            } else {
              console.error('Failed to get categories:', result.error)
              resolve([])
            }
          })
        } else {
          resolve([])
        }
      })
    },
    async getInternetHeaders() {
      return new Promise((resolve) => {
        const item = this.Office.context.mailbox.item
        if (!item || !item.getAllInternetHeadersAsync) {
          resolve({})
          return
        }

        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            // Parse headers string into key/value pairs
            const headersString = result.value
            const headers = this.parseInternetHeaders(headersString)
            resolve(headers)
          } else {
            console.error('Failed to get internet headers:', result.error)
            resolve({})
          }
        })
      })
    },
    parseInternetHeaders(headersString) {
      const headers = {}
      if (!headersString) return headers

      // Split by newlines and process each header
      const lines = headersString.split('\r\n')
      let currentKey = null
      let currentValue = ''

      for (const line of lines) {
        // Check if this is a continuation of the previous header (starts with whitespace)
        if (line.match(/^\s/) && currentKey) {
          currentValue += ' ' + line.trim()
        } else {
          // Save previous header if exists
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
    }
  }
}