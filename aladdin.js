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
      userInfo: {
        displayName: '',
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
        console.log('State saved to localStorage')
      } catch (error) {
        console.error('Error saving state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdin-state')
        if (stateJson) {
          this._state = JSON.parse(stateJson)
          console.log('State loaded from localStorage')
        }
      } catch (error) {
        console.error('Error loading state:', error)
      }
    },

    watchState() {
      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdin-state' && e.newValue) {
          try {
            this._state = JSON.parse(e.newValue)
            console.log('State updated from storage event')
            this.updateUI()
          } catch (error) {
            console.error('Error parsing storage event:', error)
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

      // Add to beginning of array (most recent first)
      this._state.events.unshift(eventEntry)

      // Keep only last 10 events
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      console.log('Event recorded:', name, details)
      this.saveState()
      this.updateUI()
    },

    async initialize() {
      console.log('Initializing Aladdin add-in')

      // Get user info
      await this.getUserInfo()

      // Register mailbox-level event handlers
      this.registerMailboxEventHandlers()

      // Register item-level event handlers if item exists
      if (this.Office.context.mailbox.item) {
        this.registerItemEventHandlers()
        await this.captureEmailDetails()
      } else {
        this.event('NoItemSelected', 'No email item is currently selected')
      }

      this.updateUI()
    },

    async getUserInfo() {
      try {
        // Get display name
        const userProfile = this.Office.context.mailbox.userProfile
        this._state.userInfo.displayName = userProfile.displayName || 'Unknown User'
        this._state.userInfo.emailAddress = userProfile.emailAddress || ''

        // Get platform
        this._state.userInfo.platform = this.Office.context.platform || 'Unknown Platform'

        // Get Office version
        this._state.userInfo.version = this.Office.context.diagnostics?.version || 'Unknown Version'

        this.saveState()
        console.log('User info captured:', this._state.userInfo)
      } catch (error) {
        console.error('Error getting user info:', error)
        this.event('UserInfoError', error.message)
      }
    },

    registerMailboxEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // ItemChanged - when user selects different item
      if (mailbox.addHandlerAsync) {
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          async (eventArgs) => {
            console.log('ItemChanged event fired')
            await this.handleItemChanged(eventArgs)
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('ItemChanged handler registered')
            } else {
              console.error('Failed to register ItemChanged handler:', result.error)
            }
          }
        )

        // OfficeThemeChanged
        mailbox.addHandlerAsync(
          this.Office.EventType.OfficeThemeChanged,
          (eventArgs) => {
            console.log('OfficeThemeChanged event fired')
            this.event('OfficeThemeChanged', { theme: eventArgs.type })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('OfficeThemeChanged handler registered')
            }
          }
        )
      }
    },

    registerItemEventHandlers() {
      const item = this.Office.context.mailbox.item

      if (!item || !item.addHandlerAsync) {
        console.log('Item or addHandlerAsync not available')
        return
      }

      // RecipientsChanged
      item.addHandlerAsync(
        this.Office.EventType.RecipientsChanged,
        async (eventArgs) => {
          console.log('RecipientsChanged event fired')
          await this.handleRecipientsChanged(eventArgs)
        },
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('RecipientsChanged handler registered')
          }
        }
      )

      // AttachmentsChanged
      item.addHandlerAsync(
        this.Office.EventType.AttachmentsChanged,
        async (eventArgs) => {
          console.log('AttachmentsChanged event fired')
          await this.handleAttachmentsChanged(eventArgs)
        },
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('AttachmentsChanged handler registered')
          }
        }
      )

      // RecurrenceChanged (for appointments)
      item.addHandlerAsync(
        this.Office.EventType.RecurrenceChanged,
        (eventArgs) => {
          console.log('RecurrenceChanged event fired')
          this.event('RecurrenceChanged', { type: eventArgs.type })
        },
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('RecurrenceChanged handler registered')
          }
        }
      )

      // AppointmentTimeChanged (for appointments)
      item.addHandlerAsync(
        this.Office.EventType.AppointmentTimeChanged,
        (eventArgs) => {
          console.log('AppointmentTimeChanged event fired')
          this.event('AppointmentTimeChanged', { type: eventArgs.type })
        },
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('AppointmentTimeChanged handler registered')
          }
        }
      )

      // EnhancedLocationsChanged (for appointments)
      item.addHandlerAsync(
        this.Office.EventType.EnhancedLocationsChanged,
        (eventArgs) => {
          console.log('EnhancedLocationsChanged event fired')
          this.event('EnhancedLocationsChanged', { type: eventArgs.type })
        },
        (result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            console.log('EnhancedLocationsChanged handler registered')
          }
        }
      )
    },

    async handleItemChanged(eventArgs) {
      const newItem = this.Office.context.mailbox.item

      // Detect deselection
      if (!newItem) {
        if (this._currentItemId) {
          this.event('ItemDeselected', {
            previousItemId: this._currentItemId
          })
          this._currentItemId = null
        }
        return
      }

      const newItemId = newItem.itemId

      // Detect selection or change
      if (newItemId !== this._currentItemId) {
        if (this._currentItemId) {
          this.event('ItemDeselected', {
            previousItemId: this._currentItemId
          })
        }

        this._currentItemId = newItemId
        this.event('ItemSelected', {
          itemId: newItemId
        })

        // Register handlers for new item
        this.registerItemEventHandlers()

        // Capture details of new item
        await this.captureEmailDetails()
      }
    },

    async handleRecipientsChanged(eventArgs) {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const recipientInfo = await this.getRecipients()
      this.event('RecipientsChanged', {
        type: eventArgs.changedRecipientFields,
        recipients: recipientInfo
      })
    },

    async handleAttachmentsChanged(eventArgs) {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const attachmentInfo = this.getAttachments()
      this.event('AttachmentsChanged', {
        type: eventArgs.attachmentStatus,
        attachmentDetails: eventArgs.attachmentDetails,
        attachments: attachmentInfo
      })
    },

    async captureEmailDetails() {
      const item = this.Office.context.mailbox.item
      if (!item) {
        console.log('No item available to capture details')
        return
      }

      try {
        const details = {
          itemId: item.itemId || 'No ID',
          itemType: item.itemType,
          subject: await this.getProperty('subject'),
          categories: item.categories || [],
          dateTimeCreated: item.dateTimeCreated?.toISOString() || null,
          dateTimeModified: item.dateTimeModified?.toISOString() || null
        }

        // Get recipients
        const recipients = await this.getRecipients()
        Object.assign(details, recipients)

        // Get attachments
        details.attachments = this.getAttachments()

        // Get internet headers
        details.internetHeaders = await this.getInternetHeaders()

        this.event('EmailDetailsCaptured', details)
      } catch (error) {
        console.error('Error capturing email details:', error)
        this.event('EmailDetailsCaptureError', error.message)
      }
    },

    async getProperty(propertyName) {
      const item = this.Office.context.mailbox.item
      if (!item) return null

      const property = item[propertyName]
      if (!property) return null

      // Check if getAsync is available
      if (typeof property.getAsync === 'function') {
        return new Promise((resolve) => {
          property.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value)
            } else {
              console.error(`Error getting ${propertyName}:`, result.error)
              resolve(null)
            }
          })
        })
      } else {
        return property
      }
    },

    async getRecipients() {
      const item = this.Office.context.mailbox.item
      if (!item) return {}

      const recipients = {}

      // Get To recipients
      if (item.to) {
        recipients.to = await this.getRecipientsAsync(item.to)
      }

      // Get From (only in read mode)
      if (item.from) {
        recipients.from = await this.getFromAsync(item.from)
      }

      // Get CC recipients
      if (item.cc) {
        recipients.cc = await this.getRecipientsAsync(item.cc)
      }

      // Get BCC recipients (compose mode only)
      if (item.bcc) {
        recipients.bcc = await this.getRecipientsAsync(item.bcc)
      }

      return recipients
    },

    async getRecipientsAsync(recipientProperty) {
      if (!recipientProperty) return []

      // Check if getAsync is available
      if (typeof recipientProperty.getAsync === 'function') {
        return new Promise((resolve) => {
          recipientProperty.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value.map(r => ({
                displayName: r.displayName,
                emailAddress: r.emailAddress
              })))
            } else {
              console.error('Error getting recipients:', result.error)
              resolve([])
            }
          })
        })
      } else if (Array.isArray(recipientProperty)) {
        return recipientProperty.map(r => ({
          displayName: r.displayName,
          emailAddress: r.emailAddress
        }))
      }

      return []
    },

    async getFromAsync(fromProperty) {
      if (!fromProperty) return null

      // Check if getAsync is available
      if (typeof fromProperty.getAsync === 'function') {
        return new Promise((resolve) => {
          fromProperty.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              const from = result.value
              resolve({
                displayName: from.displayName,
                emailAddress: from.emailAddress
              })
            } else {
              console.error('Error getting from:', result.error)
              resolve(null)
            }
          })
        })
      } else if (fromProperty.displayName || fromProperty.emailAddress) {
        return {
          displayName: fromProperty.displayName,
          emailAddress: fromProperty.emailAddress
        }
      }

      return null
    },

    getAttachments() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.attachments) return []

      return item.attachments.map(att => ({
        id: att.id,
        name: att.name,
        size: att.size,
        attachmentType: att.attachmentType,
        isInline: att.isInline || false
      }))
    },

    async getInternetHeaders() {
      const item = this.Office.context.mailbox.item
      if (!item || typeof item.getAllInternetHeadersAsync !== 'function') {
        console.log('getAllInternetHeadersAsync not available')
        return {}
      }

      return new Promise((resolve) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headersString = result.value
            const headers = this.parseInternetHeaders(headersString)
            resolve(headers)
          } else {
            console.error('Error getting internet headers:', result.error)
            resolve({})
          }
        })
      })
    },

    parseInternetHeaders(headersString) {
      if (!headersString) return {}

      const headers = {}
      const lines = headersString.split(/\r?\n/)
      let currentHeader = null
      let currentValue = ''

      for (const line of lines) {
        // Check if line starts with whitespace (continuation of previous header)
        if (line.match(/^\s/) && currentHeader) {
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
          } else {
            currentHeader = null
            currentValue = ''
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
      // Update user info section
      const userInfoEl = document.getElementById('user-info')
      if (userInfoEl) {
        const { displayName, emailAddress, platform, version } = this._state.userInfo
        userInfoEl.innerHTML = `
          <div class="user-name">${displayName}</div>
          ${emailAddress ? `<div class="user-email">${emailAddress}</div>` : ''}
          <div class="platform">Platform: ${platform}</div>
          <div class="version">Version: ${version}</div>
        `
      }

      // Update events list
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
      const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
      }
      return String(text).replace(/[&<>"']/g, m => map[m])
    }
  }
}