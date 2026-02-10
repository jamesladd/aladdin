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
      userContext: null,
      currentEmail: null
    },
    _eventHandlers: [],

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
      if (typeof window === 'undefined') return

      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdin-state' && e.newValue) {
          try {
            this._state = JSON.parse(e.newValue)
            console.log('State updated from storage:', this._state)
            this.updateUI()
          } catch (error) {
            console.error('Error parsing state from storage:', error)
          }
        }
      })
    },

    event(name, details) {
      const event = {
        name,
        details,
        timestamp: new Date().toISOString()
      }

      // Add to beginning of array and keep only last 10
      this._state.events.unshift(event)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      console.log('Event recorded:', event)
      this.saveState()
      this.updateUI()
    },

    async initialize() {
      console.log('Initializing Aladdin...')

      // Get user context
      await this.getUserContext()

      // Register all event handlers
      this.registerEventHandlers()

      // Get current email details if item exists
      if (this.Office.context.mailbox.item) {
        await this.captureEmailDetails()
      }

      // Update UI
      this.updateUI()
    },

    async getUserContext() {
      const context = this.Office.context

      this._state.userContext = {
        userName: context.mailbox?.userProfile?.displayName || 'Unknown User',
        email: context.mailbox?.userProfile?.emailAddress || '',
        platform: this.getPlatformName(context.platform),
        officeVersion: context.diagnostics?.version || 'Unknown'
      }

      this.event('UserContextLoaded', this._state.userContext)
      this.saveState()
    },

    getPlatformName(platform) {
      const Office = this.Office
      switch (platform) {
        case Office.PlatformType.PC: return 'Windows Desktop'
        case Office.PlatformType.Mac: return 'Mac Desktop'
        case Office.PlatformType.OfficeOnline: return 'Office Online'
        case Office.PlatformType.iOS: return 'iOS'
        case Office.PlatformType.Android: return 'Android'
        case Office.PlatformType.Universal: return 'Universal'
        default: return 'Unknown Platform'
      }
    },

    registerEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // Register mailbox-level events
      if (mailbox.addHandlerAsync) {
        // ItemChanged event - when user selects a different item
        this.registerMailboxEvent('ItemChanged', () => {
          this.event('ItemChanged', { message: 'User selected a different item' })
          this.captureEmailDetails()
        })

        // OfficeThemeChanged event
        this.registerMailboxEvent('OfficeThemeChanged', (args) => {
          this.event('OfficeThemeChanged', { theme: args?.type })
        })
      }

      // Register item-level events only if item exists
      if (mailbox.item && mailbox.item.addHandlerAsync) {
        // RecipientsChanged event
        this.registerItemEvent('RecipientsChanged', (args) => {
          this.event('RecipientsChanged', {
            changedRecipientFields: args?.changedRecipientFields
          })
          this.captureEmailDetails()
        })

        // AttachmentsChanged event
        this.registerItemEvent('AttachmentsChanged', (args) => {
          this.event('AttachmentsChanged', {
            attachmentStatus: args?.attachmentStatus,
            attachmentDetails: args?.attachmentDetails
          })
          this.captureEmailDetails()
        })

        // RecurrenceChanged event (for appointments)
        this.registerItemEvent('RecurrenceChanged', () => {
          this.event('RecurrenceChanged', { message: 'Recurrence pattern changed' })
        })

        // AppointmentTimeChanged event (for appointments)
        this.registerItemEvent('AppointmentTimeChanged', (args) => {
          this.event('AppointmentTimeChanged', {
            start: args?.start,
            end: args?.end
          })
        })

        // EnhancedLocationsChanged event (for appointments)
        this.registerItemEvent('EnhancedLocationsChanged', () => {
          this.event('EnhancedLocationsChanged', { message: 'Location changed' })
        })
      }

      console.log('Event handlers registered')
    },

    registerMailboxEvent(eventType, handler) {
      const mailbox = this.Office.context.mailbox

      mailbox.addHandlerAsync(
        this.Office.EventType[eventType],
        handler,
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log(`${eventType} handler registered successfully`)
          } else {
            console.warn(`Failed to register ${eventType}:`, result.error)
          }
        }
      )
    },

    registerItemEvent(eventType, handler) {
      const item = this.Office.context.mailbox.item

      item.addHandlerAsync(
        this.Office.EventType[eventType],
        handler,
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log(`${eventType} handler registered successfully`)
          } else {
            console.warn(`Failed to register ${eventType}:`, result.error)
          }
        }
      )
    },

    async captureEmailDetails() {
      const item = this.Office.context.mailbox.item

      if (!item) {
        console.log('No email item selected')
        this._state.currentEmail = null
        this.saveState()
        return
      }

      try {
        const emailDetails = {
          itemId: item.itemId,
          itemType: item.itemType,
          subject: item.subject,
          dateTimeCreated: item.dateTimeCreated,
          dateTimeModified: item.dateTimeModified,
          categories: item.categories || [],
          normalizedSubject: item.normalizedSubject,
          conversationId: item.conversationId
        }

        // Get recipients using async if available
        emailDetails.from = await this.getRecipient(item.from)
        emailDetails.to = await this.getRecipients(item.to)
        emailDetails.cc = await this.getRecipients(item.cc)
        emailDetails.bcc = await this.getRecipients(item.bcc)

        // Get attachments
        emailDetails.attachments = this.getAttachmentInfo(item.attachments)

        // Get internet headers
        emailDetails.internetHeaders = await this.getInternetHeaders()

        this._state.currentEmail = emailDetails
        this.event('EmailDetailsCaptured', {
          subject: emailDetails.subject,
          from: emailDetails.from,
          attachmentCount: emailDetails.attachments.length
        })

        this.saveState()
      } catch (error) {
        console.error('Error capturing email details:', error)
        this.event('EmailCaptureError', { error: error.message })
      }
    },

    async getRecipient(recipient) {
      if (!recipient) return null

      // Check if getAsync is available
      if (recipient.getAsync) {
        return new Promise((resolve) => {
          recipient.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value)
            } else {
              resolve({ displayName: 'Unknown', emailAddress: '' })
            }
          })
        })
      } else {
        // Use synchronous properties if available
        return {
          displayName: recipient.displayName || 'Unknown',
          emailAddress: recipient.emailAddress || ''
        }
      }
    },

    async getRecipients(recipients) {
      if (!recipients) return []

      // Check if getAsync is available
      if (recipients.getAsync) {
        return new Promise((resolve) => {
          recipients.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || [])
            } else {
              resolve([])
            }
          })
        })
      } else {
        // Return as array if it's already an array
        return Array.isArray(recipients) ? recipients : []
      }
    },

    getAttachmentInfo(attachments) {
      if (!attachments || attachments.length === 0) return []

      return attachments.map(att => ({
        id: att.id,
        name: att.name,
        attachmentType: att.attachmentType,
        size: att.size,
        contentType: att.contentType,
        isInline: att.isInline
      }))
    },

    async getInternetHeaders() {
      const item = this.Office.context.mailbox.item

      if (!item || !item.getAllInternetHeadersAsync) {
        return {}
      }

      return new Promise((resolve) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headers = this.parseInternetHeaders(result.value)
            resolve(headers)
          } else {
            console.warn('Failed to get internet headers:', result.error)
            resolve({})
          }
        })
      })
    },

    parseInternetHeaders(headerString) {
      if (!headerString) return {}

      const headers = {}
      const lines = headerString.split('\r\n')
      let currentHeader = null
      let currentValue = ''

      for (const line of lines) {
        if (line.match(/^\s/) && currentHeader) {
          // Continuation of previous header
          currentValue += ' ' + line.trim()
        } else {
          // Save previous header
          if (currentHeader) {
            headers[currentHeader] = currentValue.trim()
          }

          // Parse new header
          const colonIndex = line.indexOf(':')
          if (colonIndex > 0) {
            currentHeader = line.substring(0, colonIndex).trim()
            currentValue = line.substring(colonIndex + 1).trim()
          }
        }
      }

      // Save last header
      if (currentHeader) {
        headers[currentHeader] = currentValue.trim()
      }

      return headers
    },

    updateUI() {
      // Update user context section
      if (this._state.userContext) {
        const userNameEl = document.getElementById('user-name')
        const platformEl = document.getElementById('platform')
        const versionEl = document.getElementById('version')

        if (userNameEl) {
          userNameEl.textContent = this._state.userContext.userName
        }
        if (platformEl) {
          platformEl.textContent = `Platform: ${this._state.userContext.platform}`
        }
        if (versionEl) {
          versionEl.textContent = `Office Version: ${this._state.userContext.officeVersion}`
        }
      }

      // Update events list
      const eventsContainer = document.getElementById('events')
      if (!eventsContainer) return

      if (this._state.events.length === 0) {
        eventsContainer.innerHTML = '<div class="no-events">No events captured yet</div>'
        return
      }

      eventsContainer.innerHTML = this._state.events.map(event => `
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
    }
  }
}