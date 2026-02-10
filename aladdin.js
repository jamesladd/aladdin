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
      userInfo: {
        name: '',
        platform: '',
        version: ''
      }
    },
    _listeners: [],

    state() {
      return this._state
    },

    saveState() {
      try {
        const stateJson = JSON.stringify(this._state)
        localStorage.setItem('aladdin_state', stateJson)
        console.log('State saved:', this._state)
      } catch (e) {
        console.error('Failed to save state:', e)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdin_state')
        if (stateJson) {
          this._state = JSON.parse(stateJson)
          console.log('State loaded:', this._state)
        }
      } catch (e) {
        console.error('Failed to load state:', e)
      }
    },

    watchState() {
      // Listen to storage events from other windows/contexts
      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdin_state' && e.newValue) {
          try {
            this._state = JSON.parse(e.newValue)
            console.log('State updated from storage event:', this._state)
            this.updateUI()
          } catch (err) {
            console.error('Failed to parse storage event:', err)
          }
        }
      })
    },

    event(name, details) {
      const timestamp = new Date().toISOString()
      const event = {
        name,
        details,
        timestamp
      }

      // Add to front of array
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
      // Update status display
      const statusEl = document.getElementById('status')
      if (statusEl) {
        const { name, platform, version } = this._state.userInfo
        statusEl.innerHTML = `
          <div class="user-info">
            <div class="user-name">${name || 'Loading...'}</div>
            <div class="platform">Platform: ${platform || 'Unknown'}</div>
            <div class="platform">Version: ${version || 'Unknown'}</div>
          </div>
        `
      }

      // Update events display
      const eventsEl = document.getElementById('events')
      if (eventsEl) {
        if (this._state.events.length === 0) {
          eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        } else {
          eventsEl.innerHTML = this._state.events.map(event => `
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
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    captureEmailDetails(item) {
      if (!item) return null

      const details = {
        itemId: item.itemId,
        subject: item.subject,
        itemType: item.itemType,
        dateTimeCreated: item.dateTimeCreated,
        dateTimeModified: item.dateTimeModified
      }

      // Capture To recipients (already resolved)
      if (item.to) {
        details.to = item.to.map(recipient => ({
          displayName: recipient.displayName,
          emailAddress: recipient.emailAddress
        }))
      }

      // Capture From (already resolved)
      if (item.from) {
        details.from = {
          displayName: item.from.displayName,
          emailAddress: item.from.emailAddress
        }
      }

      // Capture CC recipients (already resolved)
      if (item.cc) {
        details.cc = item.cc.map(recipient => ({
          displayName: recipient.displayName,
          emailAddress: recipient.emailAddress
        }))
      }

      // Capture BCC recipients (already resolved) - available in compose mode
      if (item.bcc) {
        details.bcc = item.bcc.map(recipient => ({
          displayName: recipient.displayName,
          emailAddress: recipient.emailAddress
        }))
      }

      // Capture attachments
      if (item.attachments) {
        details.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          attachmentType: att.attachmentType,
          size: att.size,
          isInline: att.isInline
        }))
      }

      // Capture categories
      if (item.categories) {
        details.categories = item.categories
      }

      // Capture internet message ID
      if (item.internetMessageId) {
        details.internetMessageId = item.internetMessageId
      }

      // Capture conversation ID
      if (item.conversationId) {
        details.conversationId = item.conversationId
      }

      // Capture normalized subject
      if (item.normalizedSubject) {
        details.normalizedSubject = item.normalizedSubject
      }

      // Capture sender (read mode only)
      if (item.sender) {
        details.sender = {
          displayName: item.sender.displayName,
          emailAddress: item.sender.emailAddress
        }
      }

      return details
    },

    captureInternetHeaders(item) {
      if (!item || !item.getAllInternetHeadersAsync) {
        console.log('getAllInternetHeadersAsync not available')
        return
      }

      item.getAllInternetHeadersAsync((asyncResult) => {
        if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
          const headers = asyncResult.value
          this.event('InternetHeadersCaptured', { headers })
        } else {
          console.error('Failed to get internet headers:', asyncResult.error)
        }
      })
    },

    setupEventListeners() {
      const mailbox = this.Office.context.mailbox

      // ItemChanged event - fires when a different item is selected
      if (mailbox.addHandlerAsync) {
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          () => {
            const item = mailbox.item
            const details = this.captureEmailDetails(item)
            this.event('ItemChanged', details)

            // Also capture internet headers for the new item
            if (item) {
              this.captureInternetHeaders(item)
            }
          },
          (asyncResult) => {
            if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('ItemChanged handler registered')
            }
          }
        )
      }

      // Get current item and set up item-specific listeners
      const item = mailbox.item
      if (item) {
        // Initial capture
        const details = this.captureEmailDetails(item)
        this.event('InitialItemLoaded', details)
        this.captureInternetHeaders(item)

        // RecipientsChanged event
        if (item.addHandlerAsync) {
          item.addHandlerAsync(
            this.Office.EventType.RecipientsChanged,
            (eventArgs) => {
              const details = this.captureEmailDetails(item)
              this.event('RecipientsChanged', {
                changedRecipientFields: eventArgs.changedRecipientFields,
                ...details
              })
            },
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('RecipientsChanged handler registered')
              }
            }
          )

          // RecurrenceChanged event
          item.addHandlerAsync(
            this.Office.EventType.RecurrenceChanged,
            (eventArgs) => {
              const details = this.captureEmailDetails(item)
              this.event('RecurrenceChanged', {
                recurrence: eventArgs.recurrence,
                ...details
              })
            },
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('RecurrenceChanged handler registered')
              }
            }
          )

          // AppointmentTimeChanged event
          item.addHandlerAsync(
            this.Office.EventType.AppointmentTimeChanged,
            (eventArgs) => {
              const details = this.captureEmailDetails(item)
              this.event('AppointmentTimeChanged', {
                start: eventArgs.start,
                end: eventArgs.end,
                ...details
              })
            },
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('AppointmentTimeChanged handler registered')
              }
            }
          )

          // AttachmentsChanged event
          item.addHandlerAsync(
            this.Office.EventType.AttachmentsChanged,
            (eventArgs) => {
              const details = this.captureEmailDetails(item)
              this.event('AttachmentsChanged', {
                attachmentStatus: eventArgs.attachmentStatus,
                attachmentDetails: eventArgs.attachmentDetails,
                ...details
              })
            },
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('AttachmentsChanged handler registered')
              }
            }
          )

          // EnhancedLocationsChanged event
          item.addHandlerAsync(
            this.Office.EventType.EnhancedLocationsChanged,
            (eventArgs) => {
              const details = this.captureEmailDetails(item)
              this.event('EnhancedLocationsChanged', details)
            },
            (asyncResult) => {
              if (asyncResult.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('EnhancedLocationsChanged handler registered')
              }
            }
          )
        }
      }
    },

    getUserInfo() {
      const context = this.Office.context

      // Get user name
      if (context.mailbox && context.mailbox.userProfile) {
        this._state.userInfo.name = context.mailbox.userProfile.displayName ||
          context.mailbox.userProfile.emailAddress ||
          'Unknown User'
      }

      // Get platform
      if (context.platform) {
        const platformMap = {
          [this.Office.PlatformType.PC]: 'Windows Desktop',
          [this.Office.PlatformType.OfficeOnline]: 'Office Online',
          [this.Office.PlatformType.Mac]: 'Mac',
          [this.Office.PlatformType.iOS]: 'iOS',
          [this.Office.PlatformType.Android]: 'Android',
          [this.Office.PlatformType.Universal]: 'Universal'
        }
        this._state.userInfo.platform = platformMap[context.platform] || 'Unknown Platform'
      }

      // Get Office version
      if (context.diagnostics) {
        this._state.userInfo.version = context.diagnostics.version || 'Unknown Version'
      }

      this.saveState()
      this.updateUI()
    },

    initialize() {
      console.log('Initializing Aladdin add-in')

      // Get user information
      this.getUserInfo()

      // Set up all event listeners
      this.setupEventListeners()

      // Initial UI update
      this.updateUI()

      this.event('AddinInitialized', {
        timestamp: new Date().toISOString(),
        host: this.Office.context.host,
        platform: this._state.userInfo.platform,
        version: this._state.userInfo.version
      })
    }
  }
}