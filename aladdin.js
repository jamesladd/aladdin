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
    _itemHandlersRegistered: false,
    _state: {
      events: [],
      capturedEmail: null,
      userInfo: null,
      contactInfo: null,
      _lastCategoryInit: null
    },
    state() {
      return this._state
    },
    saveState() {
      try {
        localStorage.setItem('aladdin_state', JSON.stringify(this._state))
      } catch (e) {
        console.error('saveState error', e)
      }
    },
    loadState() {
      try {
        const raw = localStorage.getItem('aladdin_state')
        if (raw) {
          const parsed = JSON.parse(raw)
          this._state = parsed
          if (!this._state.events) this._state.events = []
        }
      } catch (e) {
        console.error('loadState error', e)
      }
    },
    watchState() {
      if (typeof window === 'undefined') return
      window.addEventListener('storage', (e) => {
        if (e.key === 'aladdin_state' && e.newValue) {
          try {
            const parsed = JSON.parse(e.newValue)
            this._state = parsed
            if (!this._state.events) this._state.events = []
            this._updateUI()
          } catch (err) {
            console.error('watchState parse error', err)
          }
        }
      })
    },
    event(name, details) {
      console.log('Event:', name)
      const entry = {
        name: name,
        timestamp: new Date().toISOString(),
        details: details || null
      }
      this._state.events.push(entry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(-10)
      }
      this.saveState()
      this._updateUI()
    },
    async initialize() {
      try {
        const mailbox = this.Office.context.mailbox

        // Rule R5: Compute userInfo synchronously and set immediately
        const userInfo = {
          userName: this._getUserName(),
          userEmail: this._getUserEmail(),
          folderName: 'Unknown',
          platform: this._getPlatform(),
          version: this._getVersion()
        }
        this._state.userInfo = userInfo
        this.saveState()
        this._updateUI()

        // Register mailbox-level events
        this._registerMailboxEvents()

        // Rule R3: Check if there is a previously captured email that needs notification
        const currentUserInfo = this._state.userInfo
        this.loadState()
        this._state.userInfo = currentUserInfo

        const previousEmail = this._state.capturedEmail
        const item = mailbox.item

        if (previousEmail) {
          let shouldNotify = false
          if (!item) {
            // No item selected
            shouldNotify = true
          } else if (!item.itemId) {
            // Compose mode (Rule R1)
            shouldNotify = true
          } else {
            // Different item
            const currentId = this._getGraphId(item.itemId)
            if (currentId !== previousEmail.graphMessageId) {
              shouldNotify = true
            }
          }
          if (shouldNotify) {
            try {
              await this.notify(previousEmail)
            } catch (e) {
              console.error('notify error during init', e)
              this._updateUI()
            }
            this._state.capturedEmail = null
            this._state.contactInfo = null
            this._currentItemId = null
            this.saveState()
          }
        }

        // Process current item
        if (item) {
          if (item.itemId) {
            // Read mode
            try {
              await this._captureCurrentItem(item)
            } catch (e) {
              console.error('captureCurrentItem error', e)
              this._updateUI()
            }
            this._registerItemEvents(item)
            try {
              await this._updateFolderFromHeaders(item)
            } catch (e) {
              console.error('updateFolderFromHeaders error', e)
              this._updateUI()
            }
            try {
              await this._fetchContactInfo(item)
            } catch (e) {
              console.error('fetchContactInfo error', e)
              this._updateUI()
            }
          } else {
            // Compose mode (Rule R1)
            this.event('ComposeMode', { status: 'composing new email' })
          }
        } else {
          this.event('NoItem', { status: 'no item selected' })
        }

        // Initialize categories (Rule R4)
        try {
          await this._initCategories()
        } catch (e) {
          console.error('initCategories error', e)
          this._updateUI()
        }
      } catch (e) {
        console.error('initialize error', e)
        this._updateUI()
      } finally {
        // Always update UI at the end
        this._updateUI()
      }
    },
    async notify(emailData) {
      if (!emailData) return
      this.event('Notify', { subject: emailData.subject, graphMessageId: emailData.graphMessageId })
      try {
        await fetch('https://www.devappeggio.com/api/inboxnotify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(emailData)
        })
      } catch (e) {
        console.error('notify fetch error', e)
      }
    },
    async getCategories(info) {
      try {
        const response = await fetch('https://www.devappeggio.com/api/inboxinit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(info)
        })
        if (response.ok) {
          const data = await response.json()
          if (Array.isArray(data)) return data
        }
      } catch (e) {
        console.error('getCategories API error, using defaults', e)
      }
      return [
        { displayName: 'InboxAgent', color: 'Preset22' }
      ]
    },
    async getContact(fromEmail) {
      if (!fromEmail) return null
      try {
        const response = await fetch('https://www.devappeggio.com/api/inboxcontact', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email: fromEmail })
        })
        if (response.ok) {
          const data = await response.json()
          return data
        }
      } catch (e) {
        console.error('getContact API error', e)
      }
      return null
    },

    // Private methods

    _getUserName() {
      try {
        return this.Office.context.mailbox.userProfile.displayName || 'Unknown'
      } catch (e) {
        return 'Unknown'
      }
    },
    _getUserEmail() {
      try {
        return this.Office.context.mailbox.userProfile.emailAddress || 'Unknown'
      } catch (e) {
        return 'Unknown'
      }
    },
    _getPlatform() {
      try {
        if (this.Office.context.diagnostics && this.Office.context.diagnostics.platform) {
          return this.Office.context.diagnostics.platform
        }
        if (this.Office.context.platform) {
          return this.Office.context.platform
        }
        return 'Unknown'
      } catch (e) {
        return 'Unknown'
      }
    },
    _getVersion() {
      try {
        if (this.Office.context.diagnostics && this.Office.context.diagnostics.version) {
          return this.Office.context.diagnostics.version
        }
        return 'Unknown'
      } catch (e) {
        return 'Unknown'
      }
    },
    _getGraphId(itemId) {
      if (!itemId) return null
      try {
        if (this.Office.context.mailbox.convertToRestId) {
          return this.Office.context.mailbox.convertToRestId(itemId, this.Office.MailboxEnums.RestVersion.v2_0)
        }
      } catch (e) {
        console.error('convertToRestId error, using raw itemId', e)
      }
      return itemId
    },
    _registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox
      const EventType = this.Office.EventType

      if (EventType.ItemChanged) {
        try {
          mailbox.addHandlerAsync(EventType.ItemChanged, (eventArgs) => {
            this._handleItemChanged(eventArgs)
          })
        } catch (e) {
          console.error('Failed to register ItemChanged', e)
        }
      }

      if (EventType.OfficeThemeChanged) {
        try {
          mailbox.addHandlerAsync(EventType.OfficeThemeChanged, (eventArgs) => {
            this.event('OfficeThemeChanged', eventArgs)
          })
        } catch (e) {
          console.error('Failed to register OfficeThemeChanged', e)
        }
      }
    },
    _registerItemEvents(item) {
      if (!item) return
      if (this._itemHandlersRegistered) return
      this._itemHandlersRegistered = true

      const EventType = this.Office.EventType
      const itemEvents = [
        'RecipientsChanged',
        'AttachmentsChanged',
        'RecurrenceChanged',
        'AppointmentTimeChanged',
        'EnhancedLocationsChanged'
      ]

      itemEvents.forEach((evtName) => {
        if (EventType[evtName]) {
          try {
            item.addHandlerAsync(EventType[evtName], (eventArgs) => {
              this.event(evtName, eventArgs)
              if (evtName === 'RecipientsChanged' || evtName === 'AttachmentsChanged') {
                this._captureCurrentItem(this.Office.context.mailbox.item)
              }
            })
          } catch (e) {
            console.error('Failed to register ' + evtName, e)
          }
        }
      })
    },
    async _handleItemChanged(eventArgs) {
      this._itemHandlersRegistered = false
      this.event('ItemChanged', { type: 'item changed' })

      // Rule R2: Re-read state from localStorage
      const currentUserInfo = this._state.userInfo
      this.loadState()
      this._state.userInfo = currentUserInfo

      const previousEmail = this._state.capturedEmail

      const item = this.Office.context.mailbox.item

      if (previousEmail) {
        let shouldNotify = false
        if (!item) {
          // No item
          shouldNotify = true
        } else if (!item.itemId) {
          // Compose mode (Rule R1)
          shouldNotify = true
        } else {
          // Different item
          const currentId = this._getGraphId(item.itemId)
          if (currentId !== previousEmail.graphMessageId) {
            shouldNotify = true
          }
        }
        if (shouldNotify) {
          // Rule R2: Re-read before notify to prevent double-notify
          const rereadUserInfo = this._state.userInfo
          this.loadState()
          this._state.userInfo = rereadUserInfo

          if (this._state.capturedEmail &&
            this._state.capturedEmail.graphMessageId === previousEmail.graphMessageId) {
            try {
              await this.notify(previousEmail)
            } catch (e) {
              console.error('notify error during item change', e)
              this._updateUI()
            }
            this._state.capturedEmail = null
            this._state.contactInfo = null
            this._currentItemId = null
            this.saveState()
          }
        }
      }

      // Process new item
      if (item) {
        if (item.itemId) {
          // Read mode
          try {
            await this._captureCurrentItem(item)
          } catch (e) {
            console.error('captureCurrentItem error', e)
            this._updateUI()
          }
          this._registerItemEvents(item)
          try {
            await this._updateFolderFromHeaders(item)
          } catch (e) {
            console.error('updateFolderFromHeaders error', e)
            this._updateUI()
          }
          try {
            await this._fetchContactInfo(item)
          } catch (e) {
            console.error('fetchContactInfo error', e)
            this._updateUI()
          }
        } else {
          // Compose mode (Rule R1)
          this.event('ComposeMode', { status: 'composing new email' })
        }
      } else {
        this.event('NoItem', { status: 'no item selected' })
      }

      this._updateUI()
    },
    async _captureCurrentItem(item) {
      if (!item) return
      if (!item.itemId) return

      const email = {
        graphMessageId: this._getGraphId(item.itemId),
        to: [],
        from: null,
        cc: [],
        subject: '',
        attachments: [],
        categories: [],
        importance: '',
        sentiment: 'neutral',
        internetHeaders: {}
      }

      // To
      try {
        email.to = await this._getRecipientsField(item, 'to')
      } catch (e) {
        console.error('Error getting To', e)
      }

      // From
      try {
        email.from = await this._getFromField(item)
      } catch (e) {
        console.error('Error getting From', e)
      }

      // CC
      try {
        email.cc = await this._getRecipientsField(item, 'cc')
      } catch (e) {
        console.error('Error getting CC', e)
      }

      // Subject
      try {
        email.subject = await this._getSubjectField(item)
      } catch (e) {
        console.error('Error getting Subject', e)
      }

      // Attachments
      try {
        if (item.attachments) {
          email.attachments = item.attachments.map(function(a) { return a.name })
        }
      } catch (e) {
        console.error('Error getting Attachments', e)
      }

      // Categories
      try {
        email.categories = await this._getCategoriesField(item)
      } catch (e) {
        console.error('Error getting Categories', e)
      }

      // Importance
      try {
        email.importance = item.importance || 'normal'
      } catch (e) {
        email.importance = 'normal'
      }

      // Internet headers
      try {
        email.internetHeaders = await this._getInternetHeaders(item)
      } catch (e) {
        console.error('Error getting Internet Headers', e)
      }

      // Sentiment from headers
      try {
        if (email.internetHeaders['X-MS-Exchange-Organization-SCL']) {
          const scl = parseInt(email.internetHeaders['X-MS-Exchange-Organization-SCL'], 10)
          if (scl >= 5) email.sentiment = 'negative'
          else if (scl >= 0) email.sentiment = 'neutral'
        }
      } catch (e) {
        console.error('Error parsing sentiment', e)
      }

      this._currentItemId = email.graphMessageId
      this._state.capturedEmail = email
      this.saveState()
      this.event('EmailCaptured', { subject: email.subject, graphMessageId: email.graphMessageId })
      this._updateUI()
    },
    async _getRecipientsField(item, fieldName) {
      const field = item[fieldName]
      if (!field) return []
      if (typeof field === 'object' && field.getAsync) {
        return new Promise((resolve) => {
          field.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(this._formatRecipients(result.value))
            } else {
              resolve([])
            }
          })
        })
      }
      if (Array.isArray(field)) {
        return this._formatRecipients(field)
      }
      return []
    },
    async _getFromField(item) {
      const from = item.from
      if (!from) return null
      if (typeof from === 'object' && from.getAsync) {
        return new Promise((resolve) => {
          from.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              const v = result.value
              resolve(v ? { name: v.displayName || '', email: v.emailAddress || '' } : null)
            } else {
              resolve(null)
            }
          })
        })
      }
      return this._formatFrom(from)
    },
    async _getSubjectField(item) {
      const subject = item.subject
      if (subject && typeof subject === 'object' && subject.getAsync) {
        return new Promise((resolve) => {
          subject.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || '')
            } else {
              resolve('')
            }
          })
        })
      }
      return typeof subject === 'string' ? subject : ''
    },
    async _getCategoriesField(item) {
      if (!item.categories) return []
      if (typeof item.categories === 'object' && item.categories.getAsync) {
        return new Promise((resolve) => {
          item.categories.getAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || [])
            } else {
              resolve([])
            }
          })
        })
      }
      if (Array.isArray(item.categories)) return item.categories
      return []
    },
    async _getInternetHeaders(item) {
      if (!item.getAllInternetHeadersAsync) return {}
      return new Promise((resolve) => {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded && result.value) {
            resolve(this._parseHeaders(result.value))
          } else {
            resolve({})
          }
        })
      })
    },
    _parseHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers
      const unfolded = headerString.replace(/\r?\n[ \t]+/g, ' ')
      const lines = unfolded.split(/\r?\n/)
      lines.forEach(function(line) {
        const idx = line.indexOf(':')
        if (idx > 0) {
          const key = line.substring(0, idx).trim()
          const value = line.substring(idx + 1).trim()
          headers[key] = value
        }
      })
      return headers
    },
    _formatRecipients(recipients) {
      if (!recipients) return []
      if (!Array.isArray(recipients)) return []
      return recipients.map(function(r) {
        return { name: r.displayName || '', email: r.emailAddress || '' }
      })
    },
    _formatFrom(from) {
      if (!from) return null
      if (from.displayName || from.emailAddress) {
        return { name: from.displayName || '', email: from.emailAddress || '' }
      }
      return null
    },
    async _updateFolderFromHeaders(item) {
      const headers = await this._getInternetHeaders(item)
      let folderName = 'Unknown'
      if (headers['X-Folder']) {
        folderName = headers['X-Folder']
      } else if (headers['X-MS-Exchange-Organization-AuthSource']) {
        folderName = headers['X-MS-Exchange-Organization-AuthSource']
      }
      try {
        if (item.parentFolderId && folderName === 'Unknown') {
          folderName = item.parentFolderId
        }
      } catch (e) {
        console.error('Error reading parentFolderId', e)
      }
      if (this._state.userInfo) {
        this._state.userInfo.folderName = folderName
        this.saveState()
        this._updateUI()
      }
    },
    async _fetchContactInfo(item) {
      if (!item) return
      const from = await this._getFromField(item)
      if (!from || !from.email) return

      const contactInfo = await this.getContact(from.email)
      if (contactInfo) {
        this._state.contactInfo = contactInfo
        this.saveState()
        this._updateUI()
      }
    },
    async _initCategories() {
      // Rule R4: Only run once every 8 hours
      const EIGHT_HOURS = 8 * 60 * 60 * 1000
      const now = Date.now()
      if (this._state._lastCategoryInit) {
        const elapsed = now - this._state._lastCategoryInit
        if (elapsed < EIGHT_HOURS) {
          return
        }
      }

      const info = this._state.userInfo || {}
      const categories = await this.getCategories(info)

      if (!categories || categories.length === 0) return

      const mapped = categories.map((cat) => {
        return {
          displayName: cat.displayName,
          color: this._mapCategoryColor(cat.color)
        }
      })

      try {
        const mailbox = this.Office.context.mailbox
        if (mailbox.masterCategories && mailbox.masterCategories.addAsync) {
          await new Promise((resolve) => {
            mailbox.masterCategories.addAsync(mapped, (result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                this.event('CategoriesAdded', { categories: mapped.map(function(c) { return c.displayName }) })
              } else {
                console.error('masterCategories.addAsync failed', result.error)
              }
              resolve()
            })
          })
        }
      } catch (e) {
        console.error('_initCategories masterCategories error', e)
      }

      this._state._lastCategoryInit = now
      this.saveState()
    },
    _mapCategoryColor(colorString) {
      if (!colorString) return 0
      try {
        const CategoryColor = this.Office.MailboxEnums.CategoryColor
        if (!CategoryColor) return colorString
        if (CategoryColor[colorString] !== undefined) {
          return CategoryColor[colorString]
        }
        const match = colorString.match(/^Preset(\d+)$/i)
        if (match) {
          const presetKey = 'Preset' + match[1]
          if (CategoryColor[presetKey] !== undefined) {
            return CategoryColor[presetKey]
          }
        }
        return CategoryColor.None || 0
      } catch (e) {
        return 0
      }
    },
    _updateUI() {
      if (typeof document === 'undefined') return

      const userNameEl = document.getElementById('userName')
      const folderNameEl = document.getElementById('folderName')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')
      const contactSection = document.getElementById('contactSection')

      const info = this._state.userInfo

      if (userNameEl) {
        if (info) {
          userNameEl.textContent = info.userName + ' (' + info.userEmail + ')'
        } else {
          userNameEl.textContent = 'Loading...'
        }
      }
      if (folderNameEl) {
        folderNameEl.textContent = (info && info.folderName) ? info.folderName : 'Unknown'
      }
      if (platformEl) {
        platformEl.textContent = (info && info.platform) ? info.platform : 'Unknown'
      }
      if (versionEl) {
        versionEl.textContent = (info && info.version) ? info.version : 'Unknown'
      }

      // Update contact section
      if (contactSection) {
        const contactInfo = this._state.contactInfo
        if (contactInfo) {
          let html = '<div class="contact-item">'
          html += '<div class="contact-name">' + (contactInfo.fullname || 'Unknown') + '</div>'
          if (contactInfo.mobile) {
            html += '<div class="contact-detail">Mobile: ' + contactInfo.mobile + '</div>'
          }
          if (contactInfo.vip) {
            html += '<div class="contact-vip">VIP</div>'
          }
          html += '</div>'
          contactSection.innerHTML = html
        } else {
          contactSection.innerHTML = '<div class="no-contact">No contact information available</div>'
        }
      }
    }
  }
}