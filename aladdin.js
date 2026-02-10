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
      items: {}, // Store item snapshots by itemId for comparison
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
          this._state = JSON.parse(stateJson)
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
              this._state = JSON.parse(e.newValue)
              console.log('State updated from storage event:', this._state)
              this.updateUI()
            } catch (error) {
              console.error('Error parsing storage event:', error)
            }
          }
        })
      }
    },

    event(name, details) {
      const event = {
        name,
        details,
        timestamp: new Date().toISOString(),
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

    async checkChanges(itemDetails) {
      if (!itemDetails || !itemDetails.itemId) return;

      const currentItemId = itemDetails.itemId
      const previousItem = this._state.items[currentItemId]

      if (!previousItem) {
        // First time seeing this item
        this._state.items[currentItemId] = itemDetails
        this.event('ItemFirstSeen', { itemId: currentItemId, subject: itemDetails.subject })
        this.saveState()
        return
      }

      // Compare all fields for changes
      const changes = []

      // Check simple fields
      const fieldsToCheck = ['subject', 'importance', 'sensitivity', 'itemClass', 'itemType']
      fieldsToCheck.forEach(field => {
        if (JSON.stringify(previousItem[field]) !== JSON.stringify(itemDetails[field])) {
          changes.push({
            field,
            oldValue: previousItem[field],
            newValue: itemDetails[field]
          })
        }
      })

      // Check recipients
      const recipientFields = ['to', 'from', 'cc', 'bcc']
      recipientFields.forEach(field => {
        const oldRecipients = JSON.stringify(previousItem[field])
        const newRecipients = JSON.stringify(itemDetails[field])
        if (oldRecipients !== newRecipients) {
          changes.push({
            field,
            oldValue: previousItem[field],
            newValue: itemDetails[field]
          })
        }
      })

      // Check categories
      if (JSON.stringify(previousItem.categories) !== JSON.stringify(itemDetails.categories)) {
        changes.push({
          field: 'categories',
          oldValue: previousItem.categories,
          newValue: itemDetails.categories
        })
      }

      // Check flags
      const flagFields = ['flagStatus', 'startDate', 'dueDate', 'completedDate']
      flagFields.forEach(field => {
        const oldFlag = previousItem.flagDetails?.[field]
        const newFlag = itemDetails.flagDetails?.[field]
        if (JSON.stringify(oldFlag) !== JSON.stringify(newFlag)) {
          changes.push({
            field: `flag.${field}`,
            oldValue: oldFlag,
            newValue: newFlag
          })
        }
      })

      // Check attachments
      if (previousItem.attachmentCount !== itemDetails.attachmentCount) {
        changes.push({
          field: 'attachmentCount',
          oldValue: previousItem.attachmentCount,
          newValue: itemDetails.attachmentCount
        })
      }

      // Check headers
      if (JSON.stringify(previousItem.internetHeaders) !== JSON.stringify(itemDetails.internetHeaders)) {
        changes.push({
          field: 'internetHeaders',
          oldValue: Object.keys(previousItem.internetHeaders || {}).length,
          newValue: Object.keys(itemDetails.internetHeaders || {}).length
        })
      }

      if (changes.length > 0) {
        this.event('ItemChanged', {
          itemId: currentItemId,
          subject: itemDetails.subject,
          changes
        })

        // Update stored item
        this._state.items[currentItemId] = itemDetails
        this.saveState()
      }
    },

    async captureItemDetails(item) {
      if (!item) return null

      const details = {
        itemId: item.itemId,
        itemClass: item.itemClass,
        itemType: item.itemType,
        subject: item.subject,
        importance: item.importance,
        sensitivity: item.sensitivity,
        categories: item.categories || [],
        attachmentCount: item.attachments?.length || 0,
        attachments: (item.attachments || []).map(a => ({
          id: a.id,
          name: a.name,
          size: a.size,
          attachmentType: a.attachmentType,
          isInline: a.isInline
        })),
        to: [],
        from: [],
        cc: [],
        bcc: [],
        internetHeaders: {},
        flagDetails: {}
      }

      // Get recipients using getAsync if available
      try {
        if (item.to && typeof item.to.getAsync === 'function') {
          details.to = await this.getRecipientsAsync(item.to)
        } else if (item.to) {
          details.to = item.to
        }

        if (item.from && typeof item.from.getAsync === 'function') {
          details.from = await this.getRecipientsAsync(item.from)
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
      }

      // Get internet headers
      if (typeof item.getAllInternetHeadersAsync === 'function') {
        try {
          details.internetHeaders = await this.getInternetHeadersAsync(item)
        } catch (error) {
          console.error('Error getting internet headers:', error)
        }
      }

      // Get flag details
      try {
        if (item.getItemIdAsync) {
          const flagStatus = await this.getPropertyAsync(item, 'getItemIdAsync')
          details.flagDetails.flagStatus = flagStatus
        }

        // Check for various flag properties
        if (item.start) details.flagDetails.startDate = item.start
        if (item.end) details.flagDetails.dueDate = item.end
      } catch (error) {
        console.error('Error getting flag details:', error)
      }

      return details
    },

    getRecipientsAsync(recipientProperty) {
      return new Promise((resolve, reject) => {
        recipientProperty.getAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded) {
            resolve(result.value.map(r => ({
              displayName: r.displayName,
              emailAddress: r.emailAddress
            })))
          } else {
            reject(result.error)
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
            reject(result.error)
          }
        })
      })
    },

    parseInternetHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers

      const lines = headerString.split(/\r?\n/)
      let currentKey = null
      let currentValue = ''

      for (let line of lines) {
        // Check if this is a continuation line (starts with whitespace)
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

    getPropertyAsync(item, methodName) {
      return new Promise((resolve, reject) => {
        if (typeof item[methodName] === 'function') {
          item[methodName]((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value)
            } else {
              reject(result.error)
            }
          })
        } else {
          resolve(null)
        }
      })
    },

    updateUI() {
      this.updateUserInfo()
      this.updateEventsList()
    },

    updateUserInfo() {
      const userNameEl = document.getElementById('user-name')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')

      if (userNameEl && this.Office?.context?.mailbox?.userProfile) {
        const profile = this.Office.context.mailbox.userProfile
        userNameEl.textContent = profile.displayName || profile.emailAddress || 'Unknown User'
      }

      if (platformEl && this.Office?.context?.platform) {
        platformEl.textContent = `Platform: ${this.Office.context.platform}`
      }

      if (versionEl && this.Office?.context?.diagnostics) {
        const version = this.Office.context.diagnostics.version || 'Unknown'
        versionEl.textContent = `Version: ${version}`
      }
    },

    updateEventsList() {
      const eventsContainer = document.getElementById('events')
      if (!eventsContainer) return

      if (this._state.events.length === 0) {
        eventsContainer.innerHTML = '<div class="no-events">No events recorded yet</div>'
        return
      }

      eventsContainer.innerHTML = this._state.events.map(event => {
        const timestamp = new Date(event.timestamp).toLocaleString()
        const detailsStr = JSON.stringify(event.details, null, 2)

        return `
          <div class="event-item">
            <div class="event-name">${this.escapeHtml(event.name)}</div>
            <div class="event-timestamp">${timestamp}</div>
            <div class="event-details">${this.escapeHtml(detailsStr)}</div>
          </div>
        `
      }).join('')
    },

    escapeHtml(text) {
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    },

    registerEventHandlers() {
      const mailbox = this.Office.context.mailbox

      // Mailbox-level events
      const mailboxEvents = [
        'ItemChanged',
        'OfficeThemeChanged'
      ]

      mailboxEvents.forEach(eventType => {
        mailbox.addHandlerAsync(
          this.Office.EventType[eventType],
          async (event) => {
            console.log(`${eventType} event fired`, event)

            if (eventType === 'ItemChanged') {
              const item = mailbox.item
              if (item && item.itemId !== this._currentItemId) {
                this._currentItemId = item.itemId
                this.event(eventType, { itemId: item.itemId, subject: item.subject })

                // Capture full item details and check for changes
                const itemDetails = await this.captureItemDetails(item)
                await this.checkChanges(itemDetails)
              }
            } else {
              this.event(eventType, { type: event.type })
            }
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log(`Successfully registered ${eventType} handler`)
            } else {
              console.error(`Failed to register ${eventType} handler:`, result.error)
            }
          }
        )
      })

      // Item-level events - only register if item exists
      const item = mailbox.item
      if (item && typeof item.addHandlerAsync === 'function') {
        const itemEvents = [
          'RecipientsChanged',
          'AttachmentsChanged',
          'RecurrenceChanged',
          'AppointmentTimeChanged',
          'EnhancedLocationsChanged'
        ]

        itemEvents.forEach(eventType => {
          // Check if this event type exists in Office.EventType
          if (this.Office.EventType[eventType]) {
            item.addHandlerAsync(
              this.Office.EventType[eventType],
              async (event) => {
                console.log(`${eventType} event fired`, event)
                this.event(eventType, { type: event.type, itemId: item.itemId })

                // Capture updated item details and check for changes
                const itemDetails = await this.captureItemDetails(item)
                await this.checkChanges(itemDetails)
              },
              (result) => {
                if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                  console.log(`Successfully registered ${eventType} handler`)
                } else {
                  console.error(`Failed to register ${eventType} handler:`, result.error)
                }
              }
            )
          }
        })
      }
    },

    async initialize() {
      console.log('Aladdin initializing...')

      // Update user info
      this.updateUserInfo()

      // Register all event handlers
      this.registerEventHandlers()

      // Capture initial item if one is selected
      const item = this.Office.context.mailbox.item
      if (item) {
        this._currentItemId = item.itemId
        this.event('AddinInitialized', { itemId: item.itemId, subject: item.subject })

        const itemDetails = await this.captureItemDetails(item)
        await this.checkChanges(itemDetails)
      } else {
        this.event('AddinInitialized', { message: 'No item selected' })
      }

      console.log('Aladdin initialized')
    },
  }
}