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
      _lastCategoryInit: null,
      showMoreContact: false
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

        // Rule R5: Compute userInfo synchronously and update UI immediately
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

        // Rule R3 & R5: Load state while preserving userInfo
        const currentUserInfo = this._state.userInfo
        this.loadState()
        this._state.userInfo = currentUserInfo
        this.watchState()

        const previousEmail = this._state.capturedEmail
        const item = mailbox.item

        // Rule R3: Check if there is a previously captured email that needs notification
        if (previousEmail) {
          let shouldNotify = false

          if (!item) {
            shouldNotify = true
          } else if (!item.itemId) {
            // Rule R1: Compose mode - explicit deselection
            shouldNotify = true
          } else {
            const currentId = this._getGraphId(item.itemId)
            if (currentId !== previousEmail.graphMessageId) {
              shouldNotify = true
            }
          }

          if (shouldNotify) {
            // Rule R2: Re-read state before notify
            const savedUserInfo = this._state.userInfo
            this.loadState()
            this._state.userInfo = savedUserInfo

            if (this._state.capturedEmail &&
              this._state.capturedEmail.graphMessageId === previousEmail.graphMessageId) {
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
            // Rule R1: Compose mode - explicit state
            this.event('ComposeMode', { status: 'composing new email' })
          }
        } else {
          this.event('NoItem', { status: 'no item selected' })
        }

        // Rule R4: Initialize categories (8-hour throttle)
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
        // Rule R5: Always update UI
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
          body: JSON.stringify({ emailAddress: emailAddress })
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
    toggleMoreContact() {
      this._state.showMoreContact = !this._state.showMoreContact
      this.saveState()
      this._updateUI()
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

      // ItemChanged - mailbox level
      if (EventType.ItemChanged && mailbox.addHandlerAsync) {
        try {
          mailbox.addHandlerAsync(EventType.ItemChanged, (eventArgs) => {
            this._handleItemChanged(eventArgs)
          })
        } catch (e) {
          console.error('Failed to register ItemChanged', e)
        }
      }

      // OfficeThemeChanged - mailbox level
      if (EventType.OfficeThemeChanged && mailbox.addHandlerAsync) {
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
        if (EventType[evtName] && item.addHandlerAsync) {
          try {
            item.addHandlerAsync(EventType[evtName], (eventArgs) => {
              this.event(evtName, eventArgs)
              if (evtName === 'RecipientsChanged' || evtName === 'AttachmentsChanged') {
                this._captureCurrentItem(this.Office.context.mailbox.item).catch(e => {
                  console.error('captureCurrentItem error in event handler', e)
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
      this._itemHandlersRegistered = false
      this.event('ItemChanged', { type: 'item changed' })

      // Rule R2: Re-read state from localStorage
      const savedUserInfo = this._state.userInfo
      this.loadState()
      this._state.userInfo = savedUserInfo

      const previousEmail = this._state.capturedEmail
      const item = this.Office.context.mailbox.item

      if (previousEmail) {
        let shouldNotify = false

        if (!item) {
          shouldNotify = true
        } else if (!item.itemId) {
          // Rule R1: Compose mode - explicit deselection
          shouldNotify = true
        } else {
          const currentId = this._getGraphId(item.itemId)
          if (currentId !== previousEmail.graphMessageId) {
            shouldNotify = true
          }
        }

        if (shouldNotify) {
          // Rule R2: Re-read before notify to prevent double-notify
          const savedUserInfo2 = this._state.userInfo
          this.loadState()
          this._state.userInfo = savedUserInfo2

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
        } else {
          // Rule R1: Compose mode
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

      // Get contact information
      if (email.from && email.from.email) {
        try {
          const contactInfo = await this.getContact(email.from.email)
          this._state.contactInfo = contactInfo
          this.saveState()
        } catch (e) {
          console.error('getContact error', e)
        }
      }

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
    _formatTimestamp(timestamp) {
      if (!timestamp) return ''
      try {
        const date = new Date(timestamp)
        if (isNaN(date.getTime())) return ''

        const day = String(date.getDate()).padStart(2, '0')
        const month = String(date.getMonth() + 1).padStart(2, '0')
        const year = date.getFullYear()
        const hours = String(date.getHours()).padStart(2, '0')
        const minutes = String(date.getMinutes()).padStart(2, '0')
        const seconds = String(date.getSeconds()).padStart(2, '0')

        return day + '/' + month + '/' + year + ' ' + hours + ':' + minutes + ':' + seconds
      } catch (e) {
        return ''
      }
    },
    _isUrl(str) {
      if (!str) return false
      try {
        const url = new URL(str)
        return url.protocol === 'http:' || url.protocol === 'https:'
      } catch (e) {
        return false
      }
    },
    _createContactField(label, value, isUrl) {
      if (!value) return ''

      let html = '<div class="contact-field">'
      html += '<span class="field-label">' + this._escapeHtml(label) + ':</span> '

      if (isUrl && this._isUrl(value)) {
        html += '<a href="' + this._escapeHtml(value) + '" target="_blank" rel="noopener noreferrer">' +
          this._escapeHtml(value) + '</a>'
      } else {
        html += '<span class="field-value">' + this._escapeHtml(value) + '</span>'
      }

      html += '</div>'

      return html
    },
    _updateUI() {
      if (typeof document === 'undefined') return

      const userNameEl = document.getElementById('userName')
      const folderNameEl = document.getElementById('folderName')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')
      const contactSectionEl = document.getElementById('contactSection')
      const actionsSectionEl = document.getElementById('actionsSection')
      const sectionTitleEl = document.querySelector('.section-title-container')

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

      // Update section title with chevron toggle
      if (sectionTitleEl) {
        const contact = this._state.contactInfo
        if (contact) {
          const chevronClass = this._state.showMoreContact ? 'chevron-up' : 'chevron-down'
          sectionTitleEl.innerHTML = '<span class="section-title">Contact</span>' +
            '<button id="toggleChevronBtn" class="chevron-btn ' + chevronClass + '" title="Toggle contact details">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
            '<polyline points="6 9 12 15 18 9"></polyline>' +
            '</svg>' +
            '</button>'
        } else {
          sectionTitleEl.innerHTML = '<span class="section-title">Contact</span>'
        }
      }

      // Update contact section
      if (contactSectionEl) {
        const contact = this._state.contactInfo
        if (contact) {
          let html = '<div class="contact-card">'

          // Primary fields
          html += '<div class="contact-name">' +
            this._escapeHtml(contact.Firstname || '') + ' ' +
            this._escapeHtml(contact.Surname || '') + '</div>'

          if (contact.VIPStatus) {
            html += '<div class="vip-badge">VIP</div>'
          }

          html += this._createContactField('Job Title', contact.JobTitle, false)
          html += this._createContactField('Company', contact.Company, false)
          html += this._createContactField('Mobile', contact.Mobile, false)
          html += this._createContactField('Account No', contact.AccountNo, false)

          // More/Less toggle
          if (this._state.showMoreContact) {
            html += '<div class="contact-more">'
            html += this._createContactField('UID', contact.UID, false)
            html += this._createContactField('Email', contact.EmailAddress, false)
            html += this._createContactField('Email Alias', contact.EmailNameAlias, false)
            html += this._createContactField('Street1', contact.Street1, false)
            html += this._createContactField('Street2', contact.Street2, false)
            html += this._createContactField('City', contact.City, false)
            html += this._createContactField('PostCode', contact.PostCode, false)

            if (contact.Linkedin) {
              html += this._createContactField('LinkedIn', contact.Linkedin, true)
            }
            if (contact.X) {
              html += this._createContactField('X', contact.X, true)
            }
            if (contact.Facebook) {
              html += this._createContactField('Facebook', contact.Facebook, true)
            }
            if (contact.Instagram) {
              html += this._createContactField('Instagram', contact.Instagram, true)
            }

            if (contact.OtherChan1) {
              html += this._createContactField('Other Channel 1', contact.OtherChan1, false)
            }
            if (contact.OtherChan2) {
              html += this._createContactField('Other Channel 2', contact.OtherChan2, false)
            }
            if (contact.SubscriberAttr1) {
              html += this._createContactField('Subscriber Attr 1', contact.SubscriberAttr1, false)
            }
            if (contact.SubscriberAttr2) {
              html += this._createContactField('Subscriber Attr 2', contact.SubscriberAttr2, false)
            }
            if (contact.SubscriberAttr3) {
              html += this._createContactField('Subscriber Attr 3', contact.SubscriberAttr3, false)
            }
            if (contact.SubscriberAttr4) {
              html += this._createContactField('Subscriber Attr 4', contact.SubscriberAttr4, false)
            }

            // Timestamp fields
            if (contact.CreatedAt) {
              html += this._createContactField('Created', this._formatTimestamp(contact.CreatedAt), false)
            }
            if (contact.UpdatedAt) {
              html += this._createContactField('Updated', this._formatTimestamp(contact.UpdatedAt), false)
            }
            if (contact.LastContactedAt) {
              html += this._createContactField('Last Contacted', this._formatTimestamp(contact.LastContactedAt), false)
            }
            html += '</div>'
            html += '<button id="toggleContactBtn" class="toggle-btn">Less</button>'
          }

          html += '</div>'
          contactSectionEl.innerHTML = html

          // Attach event listeners
          const toggleBtn = document.getElementById('toggleContactBtn')
          if (toggleBtn) {
            toggleBtn.onclick = () => {
              this.toggleMoreContact()
            }
          }

          const chevronBtn = document.getElementById('toggleChevronBtn')
          if (chevronBtn) {
            chevronBtn.onclick = () => {
              this.toggleMoreContact()
            }
          }
        } else {
          contactSectionEl.innerHTML = '<div class="no-contact">No contact information available</div>'
        }
      }

      // Update actions section
      if (actionsSectionEl) {
        actionsSectionEl.innerHTML = '<div class="no-actions">No actions available</div>'
      }
    },
    _escapeHtml(text) {
      if (!text) return ''
      const div = document.createElement('div')
      div.textContent = text
      return div.innerHTML
    }
  }
}