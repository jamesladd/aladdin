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
        userName: '',
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
        localStorage.setItem('aladdinState', stateJson)
        console.log('State saved:', this._state)
      } catch (error) {
        console.error('Error saving state:', error)
      }
    },

    loadState() {
      try {
        const stateJson = localStorage.getItem('aladdinState')
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
      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdinState' && e.newValue) {
          try {
            this._state = JSON.parse(e.newValue)
            console.log('State updated from storage event:', this._state)
            this.updateUI()
          } catch (error) {
            console.error('Error parsing storage event:', error)
          }
        }
      })
    },

    event(name, details) {
      const timestamp = new Date().toISOString()
      const eventRecord = {
        name,
        details,
        timestamp
      }

      // Add to beginning of array (most recent first)
      this._state.events.unshift(eventRecord)

      // Keep only last 10 events
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(0, 10)
      }

      console.log('Event recorded:', eventRecord)
      this.saveState()
      this.updateUI()
    },

    async captureUserInfo() {
      try {
        const context = this.Office.context

        // Get user name
        if (context.mailbox && context.mailbox.userProfile) {
          this._state.userInfo.userName = context.mailbox.userProfile.displayName ||
            context.mailbox.userProfile.emailAddress ||
            'Unknown User'
        }

        // Get platform
        this._state.userInfo.platform = this.getPlatformName(context.platform)

        // Get Office version
        this._state.userInfo.version = context.diagnostics ?
          context.diagnostics.version :
          'Unknown Version'

        console.log('User info captured:', this._state.userInfo)
        this.saveState()
        this.updateUI()
      } catch (error) {
        console.error('Error capturing user info:', error)
        this.event('Error', { message: 'Failed to capture user info', error: error.message })
      }
    },

    getPlatformName(platform) {
      const platformMap = {
        [this.Office.PlatformType.PC]: 'PC (Windows)',
        [this.Office.PlatformType.OfficeOnline]: 'Office Online',
        [this.Office.PlatformType.Mac]: 'Mac',
        [this.Office.PlatformType.iOS]: 'iOS',
        [this.Office.PlatformType.Android]: 'Android',
        [this.Office.PlatformType.Universal]: 'Universal'
      }
      return platformMap[platform] || `Unknown (${platform})`
    },

    async captureEmailDetails(item) {
      if (!item) return null

      try {
        const details = {
          itemId: item.itemId,
          subject: item.subject,
          categories: item.categories,
          conversationId: item.conversationId
        }

        // Get recipients using callbacks wrapped in promises
        const getRecipients = (recipientType) => {
          return new Promise((resolve) => {
            recipientType.getAsync((result) => {
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
        }

        // Capture To, CC, From
        if (item.to) {
          details.to = await getRecipients(item.to)
        }

        if (item.cc) {
          details.cc = await getRecipients(item.cc)
        }

        if (item.from) {
          details.from = await new Promise((resolve) => {
            item.from.getAsync((result) => {
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
        }

        // Get attachments
        if (item.attachments && item.attachments.length > 0) {
          details.attachments = item.attachments.map(att => ({
            name: att.name,
            attachmentType: att.attachmentType,
            size: att.size,
            id: att.id
          }))
        }

        // Get flags (if available)
        if (item.isRead !== undefined) {
          details.isRead = item.isRead
        }

        if (item.internetMessageId) {
          details.internetMessageId = item.internetMessageId
        }

        return details
      } catch (error) {
        console.error('Error capturing email details:', error)
        return { error: error.message }
      }
    },

    setupEventHandlers() {
      const mailbox = this.Office.context.mailbox

      if (!mailbox) {
        console.error('Mailbox not available')
        return
      }

      // Listen for item changed event (when user selects different email)
      if (mailbox.addHandlerAsync) {
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          async () => {
            console.log('ItemChanged event fired')
            this.event('ItemChanged', { timestamp: new Date().toISOString() })

            // Capture details of newly selected item
            const item = mailbox.item
            if (item) {
              const emailDetails = await this.captureEmailDetails(item)
              this.event('EmailSelected', emailDetails)
            }
          },
          (result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              console.log('ItemChanged handler registered successfully')
            } else {
              console.error('Failed to register ItemChanged handler:', result.error)
            }
          }
        )
      }

      // Capture initial item if available
      const currentItem = mailbox.item
      if (currentItem) {
        this.captureEmailDetails(currentItem).then(details => {
          this.event('InitialEmailLoaded', details)
        })
      }
    },

    updateUI() {
      // Update user info section
      const userNameEl = document.getElementById('userName')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')

      if (userNameEl) {
        userNameEl.textContent = this._state.userInfo.userName || 'Loading...'
      }
      if (platformEl) {
        platformEl.textContent = this._state.userInfo.platform || 'Loading...'
      }
      if (versionEl) {
        versionEl.textContent = this._state.userInfo.version || 'Loading...'
      }

      // Update events list
      const eventsContainer = document.getElementById('events')
      if (!eventsContainer) return

      if (this._state.events.length === 0) {
        eventsContainer.innerHTML = '<div class="no-events">No events recorded yet</div>'
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
    },

    async initialize() {
      console.log('Aladdin initializing...')

      this.event('AddinInitialized', {
        timestamp: new Date().toISOString(),
        host: this.Office.context.host
      })

      // Capture user information
      await this.captureUserInfo()

      // Setup event handlers
      this.setupEventHandlers()

      // Initial UI update
      this.updateUI()

      // Update status
      const statusEl = document.getElementById('status')
      if (statusEl) {
        statusEl.textContent = 'Ready'
        statusEl.style.color = '#4CAF50'
      }

      console.log('Aladdin initialized successfully')
    }
  }
}