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
    _previousItem: null,
    _currentItem: null,
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
        localStorage.setItem('aladdin-state', JSON.stringify(this._state))
      } catch (error) {
        console.error('Error saving state:', error)
      }
    },

    loadState() {
      try {
        const stored = localStorage.getItem('aladdin-state')
        if (stored) {
          this._state = JSON.parse(stored)
          this.updateUI()
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
            this.updateUI()
          } catch (error) {
            console.error('Error parsing storage event:', error)
          }
        }
      })
    },

    event(name, details) {
      const eventEntry = {
        name,
        details,
        timestamp: new Date().toISOString()
      }

      // Add to front, keep only last 10
      this._state.events.unshift(eventEntry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      this.saveState()
      this.updateUI()
    },

    async captureItemDetails(item) {
      if (!item) return null

      const details = {
        itemType: item.itemType,
        itemId: item.itemId,
        categories: item.categories || [],
        normalizedSubject: item.normalizedSubject || '',
        attachments: [],
        to: [],
        from: null,
        cc: [],
        subject: '',
        internetHeaders: {}
      }

      // Check if getAsync is available for async methods
      const hasGetAsync = typeof item.subject?.getAsync === 'function'

      if (hasGetAsync) {
        // Use async methods
        try {
          details.subject = await this.getAsyncValue(item.subject)
          details.to = await this.getAsyncValue(item.to)
          details.from = await this.getAsyncValue(item.from)
          details.cc = await this.getAsyncValue(item.cc)
        } catch (error) {
          console.error('Error getting async values:', error)
        }
      } else {
        // Use synchronous properties
        details.subject = item.subject || ''
        details.to = item.to || []
        details.from = item.from || null
        details.cc = item.cc || []
      }

      // Capture attachments
      if (item.attachments && item.attachments.length > 0) {
        details.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType,
          isInline: att.isInline
        }))
      }

      // Get internet headers
      if (typeof item.getAllInternetHeadersAsync === 'function') {
        try {
          const headers = await this.getAllInternetHeaders(item)
          details.internetHeaders = this.parseInternetHeaders(headers)
        } catch (error) {
          console.error('Error getting internet headers:', error)
        }
      }

      return details
    },

    getAsyncValue(property) {
      return new Promise((resolve, reject) => {
        if (!property || typeof property.getAsync !== 'function') {
          resolve(property)
          return
        }

        property.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value)
          } else {
            reject(result.error)
          }
        })
      })
    },

    getAllInternetHeaders(item) {
      return new Promise((resolve, reject) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value)
          } else {
            reject(result.error)
          }
        })
      })
    },

    parseInternetHeaders(headersString) {
      const headers = {}
      if (!headersString) return headers

      const lines = headersString.split('\r\n')
      let currentHeader = null

      for (const line of lines) {
        if (line.match(/^\s/) && currentHeader) {
          // Continuation of previous header
          headers[currentHeader] += ' ' + line.trim()
        } else {
          const colonIndex = line.indexOf(':')
          if (colonIndex > 0) {
            const key = line.substring(0, colonIndex).trim()
            const value = line.substring(colonIndex + 1).trim()
            currentHeader = key
            headers[key] = value
          }
        }
      }

      return headers
    },

    async checkChanges() {
      if (!this._previousItem || !this._currentItem) return

      const prevDetails = await this.captureItemDetails(this._previousItem)
      const currDetails = await this.captureItemDetails(this._currentItem)

      if (!prevDetails || !currDetails) return

      const changes = []

      // Check all fields
      const fields = ['subject', 'normalizedSubject', 'itemId', 'categories']
      for (const field of fields) {
        if (JSON.stringify(prevDetails[field]) !== JSON.stringify(currDetails[field])) {
          changes.push(`${field} changed`)
        }
      }

      // Check recipients
      if (JSON.stringify(prevDetails.to) !== JSON.stringify(currDetails.to)) {
        changes.push('To recipients changed')
      }
      if (JSON.stringify(prevDetails.from) !== JSON.stringify(currDetails.from)) {
        changes.push('From changed')
      }
      if (JSON.stringify(prevDetails.cc) !== JSON.stringify(currDetails.cc)) {
        changes.push('CC recipients changed')
      }

      // Check attachments
      if (JSON.stringify(prevDetails.attachments) !== JSON.stringify(currDetails.attachments)) {
        changes.push('Attachments changed')
      }

      if (changes.length > 0) {
        this.event('ItemChangesDetected', { changes, currentItem: currDetails })
      }
    },

    registerEvents() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events
      if (typeof mailbox.addHandlerAsync === 'function') {
        // ItemChanged event
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          async () => {
            const item = mailbox.item
            const details = await this.captureItemDetails(item)
            this.event('ItemChanged', details)

            this._previousItem = this._currentItem
            this._currentItem = item
            await this.checkChanges()
          }
        )

        // OfficeThemeChanged event
        mailbox.addHandlerAsync(
          this.Office.EventType.OfficeThemeChanged,
          (args) => {
            this.event('OfficeThemeChanged', args)
          }
        )
      }

      // Item-level events (only if item exists)
      if (mailbox.item && typeof mailbox.item.addHandlerAsync === 'function') {
        const item = mailbox.item

        // RecipientsChanged event
        item.addHandlerAsync(
          this.Office.EventType.RecipientsChanged,
          async (args) => {
            const details = await this.captureItemDetails(item)
            this.event('RecipientsChanged', { args, item: details })

            this._previousItem = this._currentItem
            this._currentItem = item
            await this.checkChanges()
          }
        )

        // AttachmentsChanged event
        item.addHandlerAsync(
          this.Office.EventType.AttachmentsChanged,
          async (args) => {
            const details = await this.captureItemDetails(item)
            this.event('AttachmentsChanged', { args, item: details })

            this._previousItem = this._currentItem
            this._currentItem = item
            await this.checkChanges()
          }
        )

        // RecurrenceChanged event (if appointment)
        if (item.itemType === this.Office.MailboxEnums.ItemType.Appointment) {
          item.addHandlerAsync(
            this.Office.EventType.RecurrenceChanged,
            async (args) => {
              const details = await this.captureItemDetails(item)
              this.event('RecurrenceChanged', { args, item: details })

              this._previousItem = this._currentItem
              this._currentItem = item
              await this.checkChanges()
            }
          )

          // AppointmentTimeChanged event
          item.addHandlerAsync(
            this.Office.EventType.AppointmentTimeChanged,
            async (args) => {
              const details = await this.captureItemDetails(item)
              this.event('AppointmentTimeChanged', { args, item: details })

              this._previousItem = this._currentItem
              this._currentItem = item
              await this.checkChanges()
            }
          )

          // EnhancedLocationsChanged event
          item.addHandlerAsync(
            this.Office.EventType.EnhancedLocationsChanged,
            async (args) => {
              const details = await this.captureItemDetails(item)
              this.event('EnhancedLocationsChanged', { args, item: details })

              this._previousItem = this._currentItem
              this._currentItem = item
              await this.checkChanges()
            }
          )
        }
      }
    },

    async initialize() {
      const context = this.Office.context

      // Capture user info
      if (context.mailbox && context.mailbox.userProfile) {
        this._state.userName = context.mailbox.userProfile.displayName || 'Unknown User'
      }

      // Capture platform
      if (context.diagnostics) {
        this._state.platform = context.diagnostics.platform || 'Unknown Platform'
        this._state.officeVersion = context.diagnostics.version || 'Unknown Version'
      }

      this.saveState()
      this.updateUI()

      // Register all event handlers
      this.registerEvents()

      // Capture initial item if available
      if (context.mailbox && context.mailbox.item) {
        this._currentItem = context.mailbox.item
        const details = await this.captureItemDetails(this._currentItem)
        this.event('InitialItemLoaded', details)
      }
    },

    updateUI() {
      if (typeof document === 'undefined') return

      // Update user info
      const userInfoDiv = document.getElementById('user-info')
      if (userInfoDiv) {
        userInfoDiv.innerHTML = `
          <div class="user-name">${this._state.userName}</div>
          <div class="platform">Platform: ${this._state.platform}</div>
          <div class="platform">Version: ${this._state.officeVersion}</div>
        `
      }

      // Update events list
      const eventsDiv = document.getElementById('events')
      if (eventsDiv) {
        if (this._state.events.length === 0) {
          eventsDiv.innerHTML = '<div class="no-events">No events recorded yet</div>'
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

      // Update status
      const statusDiv = document.getElementById('status')
      if (statusDiv) {
        statusDiv.textContent = 'Ready'
        statusDiv.style.color = '#28a745'
      }
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    }
  }
}