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
      userName: '',
      platform: '',
      officeVersion: '',
    },
    _listeners: [],
    _maxEvents: 10,

    state() {
      return this._state
    },

    saveState() {
      try {
        const stateJson = JSON.stringify(this._state)
        localStorage.setItem('aladdin_state', stateJson)
        localStorage.setItem('aladdin_state_timestamp', Date.now().toString())
      } catch (error) {
        console.error('Failed to save state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdin_state')
        if (stateJson) {
          this._state = JSON.parse(stateJson)
          this.updateUI()
        }
      } catch (error) {
        console.error('Failed to load state:', error)
      }
    },

    watchState() {
      // Listen for storage events from other tabs/windows
      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdin_state' && e.newValue) {
          try {
            this._state = JSON.parse(e.newValue)
            this.updateUI()
          } catch (error) {
            console.error('Failed to parse state from storage event:', error)
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

      // Add to beginning of array
      this._state.events.unshift(eventRecord)

      // Keep only last 10 events
      if (this._state.events.length > this._maxEvents) {
        this._state.events = this._state.events.slice(0, this._maxEvents)
      }

      this.saveState()
      this.updateUI()
    },

    updateUI() {
      this.displayUserInfo()
      this.displayEvents()
    },

    displayUserInfo() {
      const statusDiv = document.getElementById('status')
      if (!statusDiv) return

      const { userName, platform, officeVersion } = this._state

      statusDiv.innerHTML = `
        <div class="user-info">
          <div class="user-name">${userName || 'Loading user...'}</div>
          <div class="platform">Platform: ${platform || 'Unknown'}</div>
          <div class="platform">Office Version: ${officeVersion || 'Unknown'}</div>
        </div>
      `
    },

    displayEvents() {
      const eventsDiv = document.getElementById('events')
      if (!eventsDiv) return

      if (this._state.events.length === 0) {
        eventsDiv.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      eventsDiv.innerHTML = `
        <div class="section-title">Recent Events (Last ${this._maxEvents})</div>
        ${this._state.events.map(event => `
          <div class="event-item">
            <div class="event-name">${this.escapeHtml(event.name)}</div>
            <div class="event-timestamp">${new Date(event.timestamp).toLocaleString()}</div>
            <div class="event-details">${this.escapeHtml(JSON.stringify(event.details, null, 2))}</div>
          </div>
        `).join('')}
      `
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    async getUserInfo() {
      try {
        if (this.Office.context.mailbox.userProfile) {
          this._state.userName = this.Office.context.mailbox.userProfile.displayName ||
            this.Office.context.mailbox.userProfile.emailAddress ||
            'Unknown User'
        }

        if (this.Office.context.platform) {
          this._state.platform = this.Office.context.platform.toString()
        }

        if (this.Office.context.diagnostics) {
          this._state.officeVersion = this.Office.context.diagnostics.version || 'Unknown'
        }

        this.saveState()
        this.updateUI()
      } catch (error) {
        console.error('Failed to get user info:', error)
        this.event('Error:GetUserInfo', { error: error.message })
      }
    },

    async getEmailDetails() {
      const mailbox = this.Office.context.mailbox
      if (!mailbox || !mailbox.item) {
        return null
      }

      const item = mailbox.item
      const details = {
        subject: item.subject,
        itemId: item.itemId,
        itemType: item.itemType,
        conversationId: item.conversationId,
        categories: item.categories,
        internetMessageId: item.internetMessageId,
      }

      // Get recipients using async methods if available
      try {
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

        if (item.bcc && typeof item.bcc.getAsync === 'function') {
          details.bcc = await this.getRecipientsAsync(item.bcc)
        } else if (item.bcc) {
          details.bcc = item.bcc
        }
      } catch (error) {
        console.error('Error getting recipients:', error)
        details.recipientError = error.message
      }

      // Get attachments
      if (item.attachments && item.attachments.length > 0) {
        details.attachments = item.attachments.map(att => ({
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType,
          id: att.id,
          isInline: att.isInline
        }))
      }

      // Get internet headers
      if (typeof item.getAllInternetHeadersAsync === 'function') {
        try {
          details.internetHeaders = await this.getInternetHeadersAsync(item)
        } catch (error) {
          console.error('Error getting internet headers:', error)
          details.internetHeadersError = error.message
        }
      }

      return details
    },

    getRecipientsAsync(recipients) {
      return new Promise((resolve, reject) => {
        recipients.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value.map(r => ({
              displayName: r.displayName,
              emailAddress: r.emailAddress,
              recipientType: r.recipientType
            })))
          } else {
            reject(new Error(result.error.message))
          }
        })
      })
    },

    getRecipientAsync(recipient) {
      return new Promise((resolve, reject) => {
        recipient.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve({
              displayName: result.value.displayName,
              emailAddress: result.value.emailAddress,
              recipientType: result.value.recipientType
            })
          } else {
            reject(new Error(result.error.message))
          }
        })
      })
    },

    getInternetHeadersAsync(item) {
      return new Promise((resolve, reject) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headers = this.parseInternetHeaders(result.value)
            resolve(headers)
          } else {
            reject(new Error(result.error.message))
          }
        })
      })
    },

    parseInternetHeaders(headersString) {
      const headers = {}
      const lines = headersString.split('\r\n')
      let currentKey = null
      let currentValue = ''

      for (const line of lines) {
        // Check if line starts with whitespace (continuation of previous header)
        if (line.match(/^\s/) && currentKey) {
          currentValue += ' ' + line.trim()
        } else {
          // Save previous header if exists
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

    setupEventListeners() {
      const mailbox = this.Office.context.mailbox

      // ItemChanged event - fires when user selects a different item
      if (mailbox.addHandlerAsync) {
        this.addEventHandler(
          this.Office.EventType.ItemChanged,
          async () => {
            const details = await this.getEmailDetails()
            this.event('ItemChanged', details)
          }
        )

        // RecipientsChanged event
        this.addEventHandler(
          this.Office.EventType.RecipientsChanged,
          async () => {
            const details = await this.getEmailDetails()
            this.event('RecipientsChanged', details)
          }
        )

        // RecurrenceChanged event
        this.addEventHandler(
          this.Office.EventType.RecurrenceChanged,
          async () => {
            const details = await this.getEmailDetails()
            this.event('RecurrenceChanged', details)
          }
        )

        // AppointmentTimeChanged event
        this.addEventHandler(
          this.Office.EventType.AppointmentTimeChanged,
          async () => {
            const details = await this.getEmailDetails()
            this.event('AppointmentTimeChanged', details)
          }
        )

        // AttachmentsChanged event
        this.addEventHandler(
          this.Office.EventType.AttachmentsChanged,
          async () => {
            const details = await this.getEmailDetails()
            this.event('AttachmentsChanged', details)
          }
        )

        // EnhancedLocationsChanged event
        this.addEventHandler(
          this.Office.EventType.EnhancedLocationsChanged,
          async () => {
            const details = await this.getEmailDetails()
            this.event('EnhancedLocationsChanged', details)
          }
        )

        // OfficeThemeChanged event
        this.addEventHandler(
          this.Office.EventType.OfficeThemeChanged,
          (eventArgs) => {
            this.event('OfficeThemeChanged', eventArgs)
          }
        )
      }

      // Listen for item selection on initial load
      if (mailbox.item) {
        this.getEmailDetails().then(details => {
          this.event('InitialItemLoad', details)
        })
      }
    },

    addEventHandler(eventType, handler) {
      try {
        this.Office.context.mailbox.addHandlerAsync(
          eventType,
          handler,
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log(`Successfully added handler for ${eventType}`)
              this._listeners.push({ eventType, handler })
            } else {
              console.warn(`Failed to add handler for ${eventType}:`, result.error)
            }
          }
        )
      } catch (error) {
        console.warn(`Exception adding handler for ${eventType}:`, error)
      }
    },

    async initialize() {
      console.log('Aladdin initializing...')

      // Get user info first
      await this.getUserInfo()

      // Setup event listeners
      this.setupEventListeners()

      // Initial UI update
      this.updateUI()

      this.event('AddinInitialized', {
        timestamp: new Date().toISOString(),
        platform: this._state.platform,
        version: this._state.officeVersion
      })
    }
  }
}