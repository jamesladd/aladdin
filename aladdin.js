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
      userInfo: null,
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
        console.error('Failed to save state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdin-state')
        if (stateJson) {
          this._state = JSON.parse(stateJson)
          console.log('State loaded:', this._state)
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
              this._state = JSON.parse(e.newValue)
              console.log('State updated from storage event:', this._state)
              this.updateUI()
            } catch (error) {
              console.error('Failed to parse state from storage event:', error)
            }
          }
        })
      }
    },

    event(name, details) {
      const timestamp = new Date().toISOString()
      const event = {
        name,
        details,
        timestamp,
      }

      // Add to beginning of array (most recent first)
      this._state.events.unshift(event)

      // Keep only last 10 events
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      console.log('Event recorded:', event)
      this.saveState()
      this.updateUI()
    },

    async checkChanges() {
      if (!this._previousItem || !this._currentItem) return;

      const changes = []

      // Compare item IDs
      if (this._previousItem.itemId !== this._currentItem.itemId) {
        changes.push('Item changed')
      }

      // Compare subjects
      if (this._previousItem.subject !== this._currentItem.subject) {
        changes.push(`Subject changed from "${this._previousItem.subject}" to "${this._currentItem.subject}"`)
      }

      // Compare categories
      const prevCategories = (this._previousItem.categories || []).join(', ')
      const currCategories = (this._currentItem.categories || []).join(', ')
      if (prevCategories !== currCategories) {
        changes.push(`Categories changed from "${prevCategories}" to "${currCategories}"`)
      }

      if (changes.length > 0) {
        this.event('ItemComparison', changes.join('; '))
      }
    },

    async captureCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) {
        console.log('No item currently selected')
        return null
      }

      const itemData = {
        itemId: item.itemId,
        itemType: item.itemType,
        subject: item.subject,
        categories: item.categories || [],
        conversationId: item.conversationId,
        internetMessageId: item.internetMessageId,
        normalizedSubject: item.normalizedSubject,
      }

      // Capture recipients using async if available
      if (item.to && typeof item.to.getAsync === 'function') {
        itemData.to = await this.getRecipientsAsync(item.to)
      } else if (item.to) {
        itemData.to = item.to
      }

      if (item.from && typeof item.from.getAsync === 'function') {
        itemData.from = await this.getRecipientAsync(item.from)
      } else if (item.from) {
        itemData.from = item.from
      }

      if (item.cc && typeof item.cc.getAsync === 'function') {
        itemData.cc = await this.getRecipientsAsync(item.cc)
      } else if (item.cc) {
        itemData.cc = item.cc
      }

      // Capture attachments
      if (item.attachments) {
        itemData.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType,
          isInline: att.isInline,
        }))
      }

      // Capture internet headers
      if (item.getAllInternetHeadersAsync) {
        try {
          itemData.internetHeaders = await this.getInternetHeadersAsync(item)
        } catch (error) {
          console.error('Failed to get internet headers:', error)
          itemData.internetHeaders = {}
        }
      }

      return itemData
    },

    getRecipientsAsync(recipientsObj) {
      return new Promise((resolve) => {
        recipientsObj.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value)
          } else {
            console.error('Failed to get recipients:', result.error)
            resolve([])
          }
        })
      })
    },

    getRecipientAsync(recipientObj) {
      return new Promise((resolve) => {
        recipientObj.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value)
          } else {
            console.error('Failed to get recipient:', result.error)
            resolve(null)
          }
        })
      })
    },

    getInternetHeadersAsync(item) {
      return new Promise((resolve) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            const headers = this.parseInternetHeaders(result.value)
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
      const lines = headersString.split('\n')
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

    async handleItemChanged() {
      console.log('Item changed event triggered')

      // Set previous item before updating current
      this._previousItem = this._currentItem

      // Capture new current item
      this._currentItem = await this.captureCurrentItem()

      if (this._currentItem) {
        this.event('ItemChanged', {
          itemId: this._currentItem.itemId,
          subject: this._currentItem.subject,
          from: this._currentItem.from,
          to: this._currentItem.to,
          categories: this._currentItem.categories,
          attachmentCount: this._currentItem.attachments?.length || 0,
        })
      }

      // Check for changes between previous and current
      await this.checkChanges()
    },

    registerEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // Register mailbox-level events
      if (mailbox.addHandlerAsync) {
        // ItemChanged event
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          () => this.handleItemChanged(),
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
            } else {
              console.error('Failed to register OfficeThemeChanged handler:', result.error)
            }
          }
        )
      }

      // Register item-level events only if an item is selected
      const item = mailbox.item
      if (item && item.addHandlerAsync) {
        // RecipientsChanged event
        item.addHandlerAsync(
          this.Office.EventType.RecipientsChanged,
          async (args) => {
            const currentItem = await this.captureCurrentItem()
            this.event('RecipientsChanged', {
              changedRecipientFields: args.changedRecipientFields,
              recipients: {
                to: currentItem?.to,
                cc: currentItem?.cc,
              },
            })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('RecipientsChanged handler registered')
            } else {
              console.error('Failed to register RecipientsChanged handler:', result.error)
            }
          }
        )

        // AttachmentsChanged event
        item.addHandlerAsync(
          this.Office.EventType.AttachmentsChanged,
          async (args) => {
            const currentItem = await this.captureCurrentItem()
            this.event('AttachmentsChanged', {
              attachmentStatus: args.attachmentStatus,
              attachmentDetails: args.attachmentDetails,
              attachments: currentItem?.attachments,
            })
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('AttachmentsChanged handler registered')
            } else {
              console.error('Failed to register AttachmentsChanged handler:', result.error)
            }
          }
        )

        // RecurrenceChanged event (for appointments)
        if (item.itemType === this.Office.MailboxEnums.ItemType.Appointment) {
          item.addHandlerAsync(
            this.Office.EventType.RecurrenceChanged,
            (args) => {
              this.event('RecurrenceChanged', { recurrence: args.recurrence })
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('RecurrenceChanged handler registered')
              } else {
                console.error('Failed to register RecurrenceChanged handler:', result.error)
              }
            }
          )

          // AppointmentTimeChanged event (for appointments)
          item.addHandlerAsync(
            this.Office.EventType.AppointmentTimeChanged,
            (args) => {
              this.event('AppointmentTimeChanged', {
                start: args.start,
                end: args.end,
              })
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('AppointmentTimeChanged handler registered')
              } else {
                console.error('Failed to register AppointmentTimeChanged handler:', result.error)
              }
            }
          )

          // EnhancedLocationsChanged event (for appointments)
          item.addHandlerAsync(
            this.Office.EventType.EnhancedLocationsChanged,
            (args) => {
              this.event('EnhancedLocationsChanged', { enhancedLocations: args.enhancedLocations })
            },
            (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                console.log('EnhancedLocationsChanged handler registered')
              } else {
                console.error('Failed to register EnhancedLocationsChanged handler:', result.error)
              }
            }
          )
        }
      }
    },

    async getUserInfo() {
      const context = this.Office.context

      const userInfo = {
        displayName: context.mailbox?.userProfile?.displayName || 'Unknown User',
        platform: this.getPlatformName(context.platform),
        version: context.diagnostics?.version || 'Unknown Version',
      }

      this._state.userInfo = userInfo
      this.saveState()
      this.updateUI()

      return userInfo
    },

    getPlatformName(platform) {
      const platformMap = {
        [this.Office.PlatformType.PC]: 'PC',
        [this.Office.PlatformType.OfficeOnline]: 'OfficeOnline',
        [this.Office.PlatformType.Mac]: 'Mac',
        [this.Office.PlatformType.iOS]: 'iOS',
        [this.Office.PlatformType.Android]: 'Android',
        [this.Office.PlatformType.Universal]: 'Universal',
      }
      return platformMap[platform] || 'Unknown Platform'
    },

    updateUI() {
      // Update user info section
      const statusDiv = document.getElementById('status')
      if (statusDiv && this._state.userInfo) {
        statusDiv.innerHTML = `
          <div class="user-info">
            <div class="user-name">${this._state.userInfo.displayName}</div>
            <div class="platform">Platform: ${this._state.userInfo.platform}</div>
            <div class="platform">Version: ${this._state.userInfo.version}</div>
          </div>
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
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    async initialize() {
      console.log('Aladdin initializing...')

      // Get user info
      await this.getUserInfo()

      // Register event handlers
      this.registerEventHandlers()

      // Capture initial item if one is selected
      this._currentItem = await this.captureCurrentItem()
      if (this._currentItem) {
        this.event('AddinInitialized', {
          itemId: this._currentItem.itemId,
          subject: this._currentItem.subject,
        })
      } else {
        this.event('AddinInitialized', { message: 'No item selected' })
      }

      console.log('Aladdin initialized')
    },
  }
}