// aladdin.js

const singleton = [false]

export function createAladdin(Office) {
  if (typeof window !== 'undefined' && window.aladdinInstance) return window.aladdinInstance;
  if (singleton[0]) return singleton[0];
  const instance = aladdin(Office)
  if (typeof window !== 'undefined') window.aladdinInstance = instance;
  if (typeof window === 'undefined') singleton[0] = instance;
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
      contactInfo: null,
      userInfo: null,
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
    },
    async initialize() {
      try {
        const mailbox = this.Office.context.mailbox

        // Rule R5: Compute and set userInfo synchronously first
        const userInfo = {
          userName: this._getUserName(),
          userEmail: this._getUserEmail(),
          folderName: 'Unknown',
          platform: this._getPlatform(),
          version: this._getVersion()
        }
        this._state.userInfo = userInfo
        this._updateUI()

        // Register mailbox-level events
        this._registerMailboxEvents()

        // Rule R3 & R2: Check for previously captured email
        const currentUserInfo = this._state.userInfo
        this.loadState()
        this._state.userInfo = currentUserInfo
        this.saveState()

        const previousEmail = this._state.capturedEmail
        const item = mailbox.item

        if (previousEmail) {
          let shouldNotify = false

          // Rule R1: Explicit compose mode detection
          if (item && !item.itemId) {
            shouldNotify = true
            this.event('ComposeDetectedInit', { status: 'compose mode detected on init' })
          } else if (!item) {
            shouldNotify = true
          } else if (item.itemId) {
            const currentId = this._getGraphId(item.itemId)
            if (currentId !== previousEmail.graphMessageId) {
              shouldNotify = true
            }
          }

          if (shouldNotify) {
            // Rule R2: Re-read before notify
            this.loadState()
            if (this._state.capturedEmail &&
              this._state.capturedEmail.graphMessageId === previousEmail.graphMessageId) {
              try {
                await this.notify(previousEmail)
              } catch (e) {
                console.error('notify error during init', e)
              }
              this._state.capturedEmail = null
              this._state.contactInfo = null
              this._currentItemId = null
              this.saveState()
              this._updateUI()
            }
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
          } else {
            // Rule R1: Compose mode
            this.event('ComposeMode', { status: 'composing new email' })
            this._state.contactInfo = null
            this._updateUI()
          }
        } else {
          this.event('NoItem', { status: 'no item selected' })
          this._state.contactInfo = null
          this._updateUI()
        }

        // Rule R4: Initialize categories (8-hour throttle)
        try {
          await this._initCategories()
        } catch (e) {
          console.error('initCategories error', e)
          this._updateUI()
        }

        // Watch for state changes
        this.watchState()

      } catch (e) {
        console.error('initialize error', e)
        this._updateUI()
      } finally {
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
    async getContact(emailAddress) {
      if (!emailAddress) return null
      try {
        const response = await fetch('https://www.devappeggio.com/api/inboxcontact', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ emailAddress })
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

      // Mailbox-level events
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

      // Item-level events
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
                this._captureCurrentItem(this.Office.context.mailbox.item).catch((e) => {
                  console.error('captureCurrentItem error in event', e)
                })
              }
            })
          } catch (e) {
            console.error('Failed to register ' + evtName, e)
          }
        }
      })
    },
    async _handleItemChanged(eventArgs) {
      try {
        this._itemHandlersRegistered = false
        this.event('ItemChanged', { type: 'item changed' })

        // Rule R2: Re-read state before processing
        const currentUserInfo = this._state.userInfo
        this.loadState()
        this._state.userInfo = currentUserInfo

        const previousEmail = this._state.capturedEmail
        const item = this.Office.context.mailbox.item

        if (previousEmail) {
          let shouldNotify = false

          // Rule R1: Explicit compose mode detection
          if (item && !item.itemId) {
            shouldNotify = true
            this.event('ComposeDetectedChange', { status: 'compose mode detected on change' })
          } else if (!item) {
            shouldNotify = true
          } else if (item.itemId) {
            const currentId = this._getGraphId(item.itemId)
            if (currentId !== previousEmail.graphMessageId) {
              shouldNotify = true
            }
          }

          if (shouldNotify) {
            // Rule R2: Re-read before notify
            this.loadState()
            if (this._state.capturedEmail &&
              this._state.capturedEmail.graphMessageId === previousEmail.graphMessageId) {
              try {
                await this.notify(previousEmail)
              } catch (e) {
                console.error('notify error during item change', e)
              }
              this._state.capturedEmail = null
              this._state.contactInfo = null
              this._currentItemId = null
              this.saveState()
              this._updateUI()
            }
          }
        }

        // Process new item
        if (item) {
          if (item.itemId) {
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
          } else {
            // Rule R1: Compose mode
            this.event('ComposeMode', { status: 'composing new email' })
            this._state.contactInfo = null
            this._updateUI()
          }
        } else {
          this.event('NoItem', { status: 'no item selected' })
          this._state.contactInfo = null
          this._updateUI()
        }

      } catch (e) {
        console.error('handleItemChanged error', e)
        this._updateUI()
      } finally {
        this._updateUI()
      }
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

      // Fetch contact info
      if (email.from && email.from.email) {
        try {
          const contact = await this.getContact(email.from.email)
          this._state.contactInfo = contact
          this.saveState()
          this._updateUI()
        } catch (e) {
          console.error('Error getting contact', e)
          this._updateUI()
        }
      } else {
        this._updateUI()
      }
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
      const userEmailEl = document.getElementById('userEmail')
      const folderNameEl = document.getElementById('folderName')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')
      const contactEl = document.getElementById('contactInfo')

      const info = this._state.userInfo

      if (userNameEl) {
        userNameEl.textContent = (info && info.userName) ? info.userName : 'Loading...'
      }
      if (userEmailEl) {
        userEmailEl.textContent = (info && info.userEmail) ? info.userEmail : 'Loading...'
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

      if (contactEl) {
        this._displayContact(contactEl)
      }
    },
    _displayContact(containerEl) {
      const contact = this._state.contactInfo
      if (!contact) {
        containerEl.innerHTML = '<div class="no-contact">No contact information available</div>'
        return
      }

      let html = '<div class="contact-container">'

      // Primary fields (always visible)
      html += '<div class="contact-primary">'
      html += '<div class="contact-field"><span class="contact-label">Name:</span> ' +
        this._escapeHtml((contact.Firstname || '') + ' ' + (contact.Surname || '')) + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Mobile:</span> ' +
        this._escapeHtml(contact.Mobile || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Job Title:</span> ' +
        this._escapeHtml(contact.JobTitle || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Company:</span> ' +
        this._escapeHtml(contact.Company || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">VIP Status:</span> ' +
        (contact.Vipstatus ? 'Yes' : 'No') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Account No:</span> ' +
        this._escapeHtml(contact.AccountNo || 'N/A') + '</div>'
      html += '</div>'

      // Secondary fields (hidden by default)
      html += '<div class="contact-secondary" id="contactSecondary" style="display: none;">'
      html += '<div class="contact-field"><span class="contact-label">UID:</span> ' +
        this._escapeHtml(contact.UID || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Street 1:</span> ' +
        this._escapeHtml(contact.Street1 || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Street 2:</span> ' +
        this._escapeHtml(contact.Street2 || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">City:</span> ' +
        this._escapeHtml(contact.City || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">PostCode:</span> ' +
        this._escapeHtml(contact.PostCode || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Email:</span> ' +
        this._escapeHtml(contact.EmailAddress || 'N/A') + '</div>'
      html += '<div class="contact-field"><span class="contact-label">Email Alias:</span> ' +
        this._escapeHtml(contact.EmailNameAlias || 'N/A') + '</div>'
      if (contact.LinkedinL) {
        html += '<div class="contact-field"><span class="contact-label">LinkedIn:</span> ' +
          '<a href="' + this._escapeHtml(contact.LinkedinL) + '" target="_blank">View Profile</a></div>'
      }
      if (contact.X) {
        html += '<div class="contact-field"><span class="contact-label">X:</span> ' +
          '<a href="' + this._escapeHtml(contact.X) + '" target="_blank">View Profile</a></div>'
      }
      if (contact.Facebook) {
        html += '<div class="contact-field"><span class="contact-label">Facebook:</span> ' +
          '<a href="' + this._escapeHtml(contact.Facebook) + '" target="_blank">View Profile</a></div>'
      }
      if (contact.Instagram) {
        html += '<div class="contact-field"><span class="contact-label">Instagram:</span> ' +
          '<a href="' + this._escapeHtml(contact.Instagram) + '" target="_blank">View Profile</a></div>'
      }
      if (contact.SubscriberAttr1) {
        html += '<div class="contact-field"><span class="contact-label">Attr 1:</span> ' +
          this._escapeHtml(contact.SubscriberAttr1) + '</div>'
      }
      if (contact.SubscriberAttr2) {
        html += '<div class="contact-field"><span class="contact-label">Attr 2:</span> ' +
          this._escapeHtml(contact.SubscriberAttr2) + '</div>'
      }
      if (contact.SubscriberAttr3) {
        html += '<div class="contact-field"><span class="contact-label">Attr 3:</span> ' +
          this._escapeHtml(contact.SubscriberAttr3) + '</div>'
      }
      if (contact.SubscriberAttr4) {
        html += '<div class="contact-field"><span class="contact-label">Attr 4:</span> ' +
          this._escapeHtml(contact.SubscriberAttr4) + '</div>'
      }
      if (contact.OtherChan1) {
        html += '<div class="contact-field"><span class="contact-label">Other 1:</span> ' +
          this._escapeHtml(contact.OtherChan1) + '</div>'
      }
      if (contact.OtherChan2) {
        html += '<div class="contact-field"><span class="contact-label">Other 2:</span> ' +
          this._escapeHtml(contact.OtherChan2) + '</div>'
      }
      if (contact.CreatedAt) {
        html += '<div class="contact-field"><span class="contact-label">Created:</span> ' +
          this._formatDate(contact.CreatedAt) + '</div>'
      }
      if (contact.UpdatedAt) {
        html += '<div class="contact-field"><span class="contact-label">Updated:</span> ' +
          this._formatDate(contact.UpdatedAt) + '</div>'
      }
      if (contact.LastContactedAt) {
        html += '<div class="contact-field"><span class="contact-label">Last Contacted:</span> ' +
          this._formatDate(contact.LastContactedAt) + '</div>'
      }
      html += '</div>'

      // More button
      html += '<button class="contact-more-btn" id="contactMoreBtn">More</button>'
      html += '</div>'

      containerEl.innerHTML = html

      // Attach event listener
      const moreBtn = document.getElementById('contactMoreBtn')
      const secondaryEl = document.getElementById('contactSecondary')
      if (moreBtn && secondaryEl) {
        moreBtn.addEventListener('click', () => {
          if (secondaryEl.style.display === 'none') {
            secondaryEl.style.display = 'block'
            moreBtn.textContent = 'Less'
          } else {
            secondaryEl.style.display = 'none'
            moreBtn.textContent = 'More'
          }
        })
      }
    },
    _escapeHtml(str) {
      if (!str) return ''
      const div = document.createElement('div')
      div.textContent = str
      return div.innerHTML
    },
    _formatDate(timestamp) {
      if (!timestamp) return 'N/A'
      try {
        const date = new Date(timestamp)
        return date.toLocaleString()
      } catch (e) {
        return String(timestamp)
      }
    }
  }
}