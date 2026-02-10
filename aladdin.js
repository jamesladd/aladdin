// aladdin.js

const singleton = [false]
const MAX_EVENTS = 10
const STATE_KEY = 'aladdin_state'

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

    state() {
      return this._state
    },

    saveState() {
      try {
        const stateJson = JSON.stringify(this._state)
        localStorage.setItem(STATE_KEY, stateJson)
      } catch (err) {
        console.error('Failed to save state:', err)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem(STATE_KEY)
        if (stateJson) {
          const loadedState = JSON.parse(stateJson)
          this._state = { ...this._state, ...loadedState }
          this.updateUI()
        }
      } catch (err) {
        console.error('Failed to load state:', err)
      }
    },

    watchState() {
      // Listen to storage events to detect changes from other tabs/windows
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === STATE_KEY && e.newValue) {
            try {
              const newState = JSON.parse(e.newValue)
              this._state = { ...this._state, ...newState }
              this.updateUI()
            } catch (err) {
              console.error('Failed to parse storage event:', err)
            }
          }
        })
      }
    },

    event(name, details) {
      const timestamp = new Date().toISOString()
      const eventEntry = {
        name,
        details,
        timestamp
      }

      // Add to beginning of array and keep only last 10
      this._state.events.unshift(eventEntry)
      if (this._state.events.length > MAX_EVENTS) {
        this._state.events = this._state.events.slice(0, MAX_EVENTS)
      }

      this.saveState()
      this.updateUI()
    },

    updateUI() {
      this.updateUserInfo()
      this.updateEventsList()
    },

    updateUserInfo() {
      const userInfoDiv = document.getElementById('user-info')
      if (!userInfoDiv) return

      const { userName, platform, officeVersion } = this._state

      userInfoDiv.innerHTML = `
        <div class="user-name">${userName || 'Loading...'}</div>
        <div class="platform">Platform: ${platform || 'Unknown'}</div>
        <div class="platform">Office Version: ${officeVersion || 'Unknown'}</div>
      `
    },

    updateEventsList() {
      const eventsDiv = document.getElementById('events')
      if (!eventsDiv) return

      if (this._state.events.length === 0) {
        eventsDiv.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      eventsDiv.innerHTML = this._state.events.map(evt => `
        <div class="event-item">
          <div class="event-name">${this.escapeHtml(evt.name)}</div>
          <div class="event-timestamp">${new Date(evt.timestamp).toLocaleString()}</div>
          <div class="event-details">${this.escapeHtml(JSON.stringify(evt.details, null, 2))}</div>
        </div>
      `).join('')
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    async getUserContext() {
      try {
        const Office = this.Office

        // Get user name
        if (Office.context.mailbox?.userProfile?.displayName) {
          this._state.userName = Office.context.mailbox.userProfile.displayName
        }

        // Get platform
        if (Office.context.platform) {
          const platformMap = {
            [Office.PlatformType.PC]: 'Windows Desktop',
            [Office.PlatformType.OfficeOnline]: 'Office Online',
            [Office.PlatformType.Mac]: 'Mac',
            [Office.PlatformType.iOS]: 'iOS',
            [Office.PlatformType.Android]: 'Android',
            [Office.PlatformType.Universal]: 'Universal'
          }
          this._state.platform = platformMap[Office.context.platform] || 'Unknown'
        }

        // Get Office version
        if (Office.context.diagnostics?.version) {
          this._state.officeVersion = Office.context.diagnostics.version
        }

        this.saveState()
        this.updateUI()
      } catch (err) {
        console.error('Failed to get user context:', err)
        this.event('Error', { message: 'Failed to get user context', error: err.message })
      }
    },

    async captureEmailItem() {
      try {
        const Office = this.Office
        const item = Office.context.mailbox?.item

        if (!item) {
          console.log('No item selected')
          return
        }

        const emailData = {
          subject: item.subject,
          itemType: item.itemType,
          itemId: item.itemId,
          conversationId: item.conversationId,
          categories: item.categories,
          dateTimeCreated: item.dateTimeCreated?.toISOString(),
          dateTimeModified: item.dateTimeModified?.toISOString(),
        }

        // Get To recipients
        if (item.to) {
          emailData.to = await this.getRecipients(item.to)
        }

        // Get From (for received items)
        if (item.from) {
          emailData.from = await this.getRecipient(item.from)
        }

        // Get Sender (for received items)
        if (item.sender) {
          emailData.sender = await this.getRecipient(item.sender)
        }

        // Get CC recipients
        if (item.cc) {
          emailData.cc = await this.getRecipients(item.cc)
        }

        // Get attachments
        if (item.attachments && item.attachments.length > 0) {
          emailData.attachments = item.attachments.map(att => ({
            id: att.id,
            name: att.name,
            size: att.size,
            attachmentType: att.attachmentType,
            isInline: att.isInline,
            contentType: att.contentType
          }))
        }

        // Get internet headers if available
        if (item.getAllInternetHeadersAsync) {
          emailData.internetHeaders = await this.getInternetHeaders(item)
        }

        this.event('EmailItemCaptured', emailData)
      } catch (err) {
        console.error('Failed to capture email item:', err)
        this.event('Error', { message: 'Failed to capture email item', error: err.message })
      }
    },

    async getRecipients(recipientsObj) {
      return new Promise((resolve) => {
        // Check if getAsync is available
        if (typeof recipientsObj.getAsync === 'function') {
          recipientsObj.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value.map(r => ({
                displayName: r.displayName,
                emailAddress: r.emailAddress,
                recipientType: r.recipientType
              })))
            } else {
              resolve([])
            }
          })
        } else {
          // Fallback for synchronous access
          resolve([])
        }
      })
    },

    async getRecipient(recipientObj) {
      return new Promise((resolve) => {
        // Check if getAsync is available
        if (typeof recipientObj.getAsync === 'function') {
          recipientObj.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              const recipient = result.value
              resolve({
                displayName: recipient.displayName,
                emailAddress: recipient.emailAddress
              })
            } else {
              resolve(null)
            }
          })
        } else {
          // Fallback for synchronous access
          resolve({
            displayName: recipientObj.displayName || '',
            emailAddress: recipientObj.emailAddress || ''
          })
        }
      })
    },

    async getInternetHeaders(item) {
      return new Promise((resolve) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headersString = result.value
            const headersParsed = this.parseInternetHeaders(headersString)
            resolve(headersParsed)
          } else {
            resolve({})
          }
        })
      })
    },

    parseInternetHeaders(headersString) {
      const headers = {}
      const lines = headersString.split(/\r?\n/)
      let currentKey = null
      let currentValue = ''

      for (let line of lines) {
        // Check if this is a continuation of the previous header (starts with whitespace)
        if (line.match(/^\s/) && currentKey) {
          currentValue += ' ' + line.trim()
        } else {
          // Save previous header if exists
          if (currentKey) {
            if (headers[currentKey]) {
              // If header already exists, convert to array
              if (Array.isArray(headers[currentKey])) {
                headers[currentKey].push(currentValue)
              } else {
                headers[currentKey] = [headers[currentKey], currentValue]
              }
            } else {
              headers[currentKey] = currentValue
            }
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
        if (headers[currentKey]) {
          if (Array.isArray(headers[currentKey])) {
            headers[currentKey].push(currentValue)
          } else {
            headers[currentKey] = [headers[currentKey], currentValue]
          }
        } else {
          headers[currentKey] = currentValue
        }
      }

      return headers
    },

    registerEvents() {
      const Office = this.Office

      // Register mailbox-level events
      if (Office.context.mailbox?.addHandlerAsync) {
        // ItemChanged - fires when a different item is selected
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          () => {
            this.event('ItemChanged', { timestamp: new Date().toISOString() })
            this.captureEmailItem()
          }
        )

        // OfficeThemeChanged - fires when the Office theme changes
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.OfficeThemeChanged,
          (args) => {
            this.event('OfficeThemeChanged', args)
          }
        )
      }

      // Register item-level events only if item exists
      if (Office.context.mailbox?.item?.addHandlerAsync) {
        const item = Office.context.mailbox.item

        // RecipientsChanged - fires when recipients are changed
        item.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          (args) => {
            this.event('RecipientsChanged', args)
            this.captureEmailItem()
          }
        )

        // AttachmentsChanged - fires when attachments are added or removed
        item.addHandlerAsync(
          Office.EventType.AttachmentsChanged,
          (args) => {
            this.event('AttachmentsChanged', args)
            this.captureEmailItem()
          }
        )

        // RecurrenceChanged - fires when recurrence pattern changes (appointments)
        item.addHandlerAsync(
          Office.EventType.RecurrenceChanged,
          (args) => {
            this.event('RecurrenceChanged', args)
            this.captureEmailItem()
          }
        )

        // AppointmentTimeChanged - fires when appointment time changes
        item.addHandlerAsync(
          Office.EventType.AppointmentTimeChanged,
          (args) => {
            this.event('AppointmentTimeChanged', args)
            this.captureEmailItem()
          }
        )

        // EnhancedLocationsChanged - fires when enhanced locations change
        item.addHandlerAsync(
          Office.EventType.EnhancedLocationsChanged,
          (args) => {
            this.event('EnhancedLocationsChanged', args)
            this.captureEmailItem()
          }
        )
      }
    },

    async initialize() {
      try {
        this.event('AddinInitialized', { timestamp: new Date().toISOString() })

        await this.getUserContext()
        this.registerEvents()
        await this.captureEmailItem()

        this.updateUI()
      } catch (err) {
        console.error('Initialization failed:', err)
        this.event('Error', { message: 'Initialization failed', error: err.message })
      }
    }
  }
}