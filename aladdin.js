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
    },
    _eventHandlersRegistered: false,

    state() {
      return this._state
    },

    saveState() {
      try {
        const stateJson = JSON.stringify(this._state)
        localStorage.setItem('aladdinState', stateJson)
        localStorage.setItem('aladdinStateTimestamp', Date.now().toString())
      } catch (error) {
        console.error('Error saving state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdinState')
        if (stateJson) {
          this._state = JSON.parse(stateJson)
          this.updateUI()
        }
      } catch (error) {
        console.error('Error loading state:', error)
      }
    },

    watchState() {
      // Listen for storage events from other windows/tabs
      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdinState' && e.newValue) {
          try {
            this._state = JSON.parse(e.newValue)
            this.updateUI()
          } catch (error) {
            console.error('Error parsing storage event:', error)
          }
        }
      })
    },

    event(name, details) {
      const eventRecord = {
        name,
        details,
        timestamp: new Date().toISOString()
      }

      // Add to beginning of array and keep only last 10
      this._state.events.unshift(eventRecord)
      this._state.events = this._state.events.slice(0, 10)

      this.saveState()
      this.updateUI()
    },

    updateUI() {
      // Update user info section
      const userInfoEl = document.getElementById('user-info')
      if (userInfoEl && this._state.userInfo) {
        const { displayName, platform, version } = this._state.userInfo
        userInfoEl.innerHTML = `
          <div class="user-name">${displayName || 'Unknown User'}</div>
          <div class="platform">Platform: ${platform || 'Unknown'}</div>
          <div class="version">Version: ${version || 'Unknown'}</div>
        `
      }

      // Update events list
      const eventsEl = document.getElementById('events')
      if (eventsEl) {
        if (this._state.events.length === 0) {
          eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        } else {
          eventsEl.innerHTML = this._state.events.map(evt => `
            <div class="event-item">
              <div class="event-name">${this.escapeHtml(evt.name)}</div>
              <div class="event-timestamp">${new Date(evt.timestamp).toLocaleString()}</div>
              <div class="event-details">${this.escapeHtml(JSON.stringify(evt.details, null, 2))}</div>
            </div>
          `).join('')
        }
      }

      // Update status
      const statusEl = document.getElementById('status')
      if (statusEl) {
        statusEl.textContent = 'Ready'
      }
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    async getUserInfo() {
      try {
        const mailbox = this.Office.context.mailbox
        const displayName = mailbox.userProfile?.displayName || 'Unknown User'
        const platform = this.Office.context.platform || 'Unknown Platform'
        const version = this.Office.context.diagnostics?.version || 'Unknown Version'

        this._state.userInfo = { displayName, platform, version }
        this.saveState()
        this.updateUI()
      } catch (error) {
        console.error('Error getting user info:', error)
        this.event('Error', { error: 'Failed to get user info', message: error.message })
      }
    },

    async getEmailItemDetails() {
      try {
        const item = this.Office.context.mailbox.item
        if (!item) {
          return { error: 'No item selected' }
        }

        const details = {
          itemType: item.itemType,
          itemId: item.itemId,
          subject: item.subject,
          categories: item.categories,
        }

        // Get recipients using getAsync if available
        if (item.to && typeof item.to.getAsync === 'function') {
          details.to = await this.getRecipientsAsync(item.to)
        } else if (item.to) {
          details.to = item.to
        }

        if (item.from && typeof item.from.getAsync === 'function') {
          details.from = await this.getRecipientAsync(item.from)
        } else if (item.from) {
          details.from = item.from
        }

        if (item.cc && typeof item.cc.getAsync === 'function') {
          details.cc = await this.getRecipientsAsync(item.cc)
        } else if (item.cc) {
          details.cc = item.cc
        }

        // Get attachments
        if (item.attachments) {
          details.attachments = item.attachments.map(att => ({
            id: att.id,
            name: att.name,
            size: att.size,
            attachmentType: att.attachmentType,
            isInline: att.isInline
          }))
        }

        // Get flags and other properties
        details.conversationId = item.conversationId
        details.dateTimeCreated = item.dateTimeCreated
        details.dateTimeModified = item.dateTimeModified
        details.importance = item.importance
        details.internetMessageId = item.internetMessageId
        details.normalizedSubject = item.normalizedSubject

        // Get internet headers
        if (typeof item.getAllInternetHeadersAsync === 'function') {
          details.internetHeaders = await this.getInternetHeaders()
        }

        return details
      } catch (error) {
        console.error('Error getting email item details:', error)
        return { error: error.message }
      }
    },

    getRecipientsAsync(recipients) {
      return new Promise((resolve) => {
        recipients.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value.map(r => ({
              displayName: r.displayName,
              emailAddress: r.emailAddress
            })))
          } else {
            resolve([])
          }
        })
      })
    },

    getRecipientAsync(recipient) {
      return new Promise((resolve) => {
        recipient.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve({
              displayName: result.value.displayName,
              emailAddress: result.value.emailAddress
            })
          } else {
            resolve(null)
          }
        })
      })
    },

    getInternetHeaders() {
      return new Promise((resolve) => {
        this.Office.context.mailbox.item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headers = this.parseInternetHeaders(result.value)
            resolve(headers)
          } else {
            resolve(null)
          }
        })
      })
    },

    parseInternetHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers

      const lines = headerString.split('\r\n')
      let currentKey = null

      for (const line of lines) {
        if (line.match(/^\s/) && currentKey) {
          // Continuation of previous header
          headers[currentKey] += ' ' + line.trim()
        } else {
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

    registerEventHandlers() {
      if (this._eventHandlersRegistered) return

      const mailbox = this.Office.context.mailbox

      // Register mailbox-level events
      if (typeof mailbox.addHandlerAsync === 'function') {
        // ItemChanged event - when user selects a different item
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          async () => {
            const details = await this.getEmailItemDetails()
            this.event('ItemChanged', details)
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('ItemChanged handler registered')
            } else {
              console.error('Failed to register ItemChanged handler:', result.error)
            }
          }
        )

        // OfficeThemeChanged event
        mailbox.addHandlerAsync(
          this.Office.EventType.OfficeThemeChanged,
          (args) => {
            this.event('OfficeThemeChanged', { theme: args })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('OfficeThemeChanged handler registered')
            }
          }
        )
      }

      // Register item-level events only if item exists
      const item = mailbox.item
      if (item && typeof item.addHandlerAsync === 'function') {
        // RecipientsChanged event
        item.addHandlerAsync(
          this.Office.EventType.RecipientsChanged,
          async (args) => {
            const details = await this.getEmailItemDetails()
            this.event('RecipientsChanged', { changedRecipientFields: args.changedRecipientFields, ...details })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('RecipientsChanged handler registered')
            }
          }
        )

        // AttachmentsChanged event
        item.addHandlerAsync(
          this.Office.EventType.AttachmentsChanged,
          async (args) => {
            const details = await this.getEmailItemDetails()
            this.event('AttachmentsChanged', { attachmentStatus: args.attachmentStatus, attachmentDetails: args.attachmentDetails, ...details })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('AttachmentsChanged handler registered')
            }
          }
        )

        // RecurrenceChanged event (for appointments)
        if (item.itemType === this.Office.MailboxEnums.ItemType.Appointment) {
          item.addHandlerAsync(
            this.Office.EventType.RecurrenceChanged,
            async () => {
              const details = await this.getEmailItemDetails()
              this.event('RecurrenceChanged', details)
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('RecurrenceChanged handler registered')
              }
            }
          )

          // AppointmentTimeChanged event
          item.addHandlerAsync(
            this.Office.EventType.AppointmentTimeChanged,
            async (args) => {
              const details = await this.getEmailItemDetails()
              this.event('AppointmentTimeChanged', { start: args.start, end: args.end, ...details })
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('AppointmentTimeChanged handler registered')
              }
            }
          )

          // EnhancedLocationsChanged event
          item.addHandlerAsync(
            this.Office.EventType.EnhancedLocationsChanged,
            async () => {
              const details = await this.getEmailItemDetails()
              this.event('EnhancedLocationsChanged', details)
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('EnhancedLocationsChanged handler registered')
              }
            }
          )
        }
      }

      this._eventHandlersRegistered = true
    },

    async initialize() {
      console.log('Aladdin initializing...')

      // Get and display user info
      await this.getUserInfo()

      // Register all event handlers
      this.registerEventHandlers()

      // Get initial email item details if available
      if (this.Office.context.mailbox.item) {
        const details = await this.getEmailItemDetails()
        this.event('Initialize', { message: 'Add-in initialized', itemDetails: details })
      } else {
        this.event('Initialize', { message: 'Add-in initialized', note: 'No item selected' })
      }

      console.log('Aladdin initialized')
    }
  }
}