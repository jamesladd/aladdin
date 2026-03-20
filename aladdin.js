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
  console.log('Aladdin version: 1.90.0', new Date());
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
      showMoreContact: false,
      isEditingContact: false,
      editedContact: null
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

            // Preserve current edit state to prevent race conditions
            const currentEditState = this._state.isEditingContact
            const currentEditedContact = this._state.editedContact

            this._state = parsed
            if (!this._state.events) this._state.events = []

            // If we're currently in edit mode, don't let storage events override it
            if (currentEditState) {
              this._state.isEditingContact = currentEditState
              this._state.editedContact = currentEditedContact
            }

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

        console.log('Mailbox diagnostics:', {
          platform: this._getPlatform(),
          itemType: mailbox.item ? mailbox.item.itemType : 'none',
          itemMode: mailbox.item ? (mailbox.item.itemId ? 'read' : 'compose') : 'none'
        })

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
            const currentId = item.itemId
            if (currentId !== previousEmail.itemId) {
              shouldNotify = true
            }
          }

          if (shouldNotify) {
            // Rule R2: Re-read state before notify
            const savedUserInfo = this._state.userInfo
            this.loadState()
            this._state.userInfo = savedUserInfo

            if (this._state.capturedEmail &&
              this._state.capturedEmail.itemId === previousEmail.itemId) {
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

      // Skip if offline
      if (typeof navigator !== 'undefined' && !navigator.onLine) {
        this.event('NotifySkipped', { reason: 'offline' })
        return
      }

      this.event('Notify', { subject: emailData.subject, itemId: emailData.itemId })

      try {
        const controller = new AbortController()
        const timeoutId = setTimeout(() => controller.abort(), 5000)
        const response = await fetch('https://www.devappeggio.com/api/inboxnotify', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify(emailData),
          signal: controller.signal,
          mode: 'cors'
        })

        clearTimeout(timeoutId)
        if (response.ok) {
          this.event('NotifySuccess', { itemId: emailData.itemId })
        } else {
          this.event('NotifyFailed', { status: response.status })
          console.warn('notify returned status:', response.status)
        }
      } catch (e) {
        if (e.name === 'AbortError') {
          this.event('NotifyTimeout', { itemId: emailData.itemId })
          console.warn('notify timed out after 5s')
        } else {
          this.event('NotifyError', { error: e.message })
          console.warn('notify error:', e.message)
        }
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
          body: JSON.stringify({
            emailAddress: emailAddress
          })
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
    async saveContact(contactData) {
      if (!contactData) return false
      this.event('SaveContact', { uid: contactData.UID })
      try {
        const response = await fetch('https://www.devappeggio.com/api/inboxcontact/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(contactData)
        })

        // Consider both 200 and 204 as success
        if (response.ok) {
          return true
        }

        console.error('saveContact failed with status:', response.status)
        return false
      } catch (e) {
        console.error('saveContact API error', e)
        return false
      }
    },
    toggleMoreContact() {
      this._state.showMoreContact = !this._state.showMoreContact
      this.saveState()
      this._updateUI()
    },
    startEditingContact() {
      if (!this._state.contactInfo) return
      this._state.isEditingContact = true
      this._state.showMoreContact = true
      this._state.editedContact = JSON.parse(JSON.stringify(this._state.contactInfo))
      this.saveState()
      this._updateUI()
    },
    cancelEditingContact() {
      this._state.isEditingContact = false
      this._state.editedContact = null
      this.saveState()
      this._updateUI()
    },
    updateEditedContactField(fieldName, value) {
      if (!this._state.editedContact) return
      this._state.editedContact[fieldName] = value
      this.saveState()
    },
    updateEditedContactVIPStatus(checked) {
      if (!this._state.editedContact) return
      this._state.editedContact.VIPStatus = checked
      this.saveState()
    },
    async saveEditedContact() {
      if (!this._state.editedContact) return

      const saveBtn = document.getElementById('saveContactBtn')
      const cancelBtn = document.getElementById('cancelContactBtn')

      if (saveBtn) {
        saveBtn.disabled = true
        saveBtn.textContent = 'Saving...'
      }
      if (cancelBtn) {
        cancelBtn.disabled = true
      }

      try {
        const success = await this.saveContact(this._state.editedContact)

        // Update contact info if successful
        if (success) {
          this._state.contactInfo = JSON.parse(JSON.stringify(this._state.editedContact))
        }

        // ALWAYS exit edit mode after save attempt - do this BEFORE saveState
        this._state.isEditingContact = false
        this._state.editedContact = null

        // Save state and update UI
        this.saveState()
        this._updateUI()

        // Show error message if save failed
        if (!success) {
          this._showAlert('Failed to save contact. Changes were not persisted to the server.')
        }
      } catch (e) {
        console.error('saveEditedContact error', e)

        // Exit edit mode even on error
        this._state.isEditingContact = false
        this._state.editedContact = null
        this.saveState()
        this._updateUI()

        this._showAlert('Error saving contact. Please try again.')
      }
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
    _registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox
      const EventType = this.Office.EventType

      // ItemChanged - mailbox level
      if (EventType.ItemChanged && mailbox.addHandlerAsync) {
        try {
          mailbox.addHandlerAsync(EventType.ItemChanged, (eventArgs) => {
            // Wrap in try-catch to prevent propagation
            try {
              this._handleItemChanged(eventArgs)
            } catch (e) {
              console.error('Error in ItemChanged handler:', e)
            }
          })
        } catch (e) {
          console.error('Failed to register ItemChanged', e)
        }
      }

      // OfficeThemeChanged - mailbox level
      if (EventType.OfficeThemeChanged && mailbox.addHandlerAsync) {
        try {
          mailbox.addHandlerAsync(EventType.OfficeThemeChanged, (eventArgs) => {
            try {
              this.event('OfficeThemeChanged', eventArgs)
            } catch (e) {
              console.error('Error in OfficeThemeChanged handler:', e)
            }
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
              try {
                this.event(evtName, eventArgs)
                if (evtName === 'RecipientsChanged' || evtName === 'AttachmentsChanged') {
                  this._captureCurrentItem(this.Office.context.mailbox.item).catch(e => {
                    console.error('captureCurrentItem error in event handler', e)
                  })
                }
              } catch (e) {
                console.error('Error in ' + evtName + ' handler:', e)
              }
            })
          } catch (e) {
            console.warn('Failed to register ' + evtName + ' (may not be supported in shared mailbox):', e)
          }
        }
      })
    },
    async _handleItemChanged(eventArgs) {
      this._itemHandlersRegistered = false
      this.event('ItemChanged', { type: 'item changed' })

      // Check if we're in edit mode with unsaved changes
      if (this._state.isEditingContact) {
        const hasEdits = this._hasContactEdits()

        if (hasEdits) {
          // Prompt user to save or discard changes
          const shouldSave = await this._showConfirmDialog(
            'Unsaved Changes',
            'You have unsaved contact changes. Do you want to save them?',
            'Save',
            'Discard'
          )

          if (shouldSave) {
            // Save changes before proceeding
            await this.saveEditedContact()
          } else {
            // Discard changes
            this._state.isEditingContact = false
            this._state.editedContact = null
            this.saveState()
          }
        } else {
          // No edits, silently exit edit mode
          this._state.isEditingContact = false
          this._state.editedContact = null
          this.saveState()
        }
      }

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
          const currentId = item.itemId
          if (currentId !== previousEmail.itemId) {
            shouldNotify = true
          }
        }

        if (shouldNotify) {
          // Rule R2: Re-read before notify to prevent double-notify
          const savedUserInfo2 = this._state.userInfo
          this.loadState()
          this._state.userInfo = savedUserInfo2

          if (this._state.capturedEmail &&
            this._state.capturedEmail.itemId === previousEmail.itemId) {
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

      // Check if we're in edit mode - this handles the case when item selection
      // triggers capture without going through _handleItemChanged
      if (this._state.isEditingContact) {
        const hasEdits = this._hasContactEdits()

        if (hasEdits) {
          // Prompt user to save or discard changes
          const shouldSave = await this._showConfirmDialog(
            'Unsaved Changes',
            'You have unsaved contact changes. Do you want to save them?',
            'Save',
            'Discard'
          )

          if (shouldSave) {
            // Save changes before proceeding
            await this.saveEditedContact()
          } else {
            // Discard changes
            this._state.isEditingContact = false
            this._state.editedContact = null
            this.saveState()
          }
        } else {
          // No edits, silently exit edit mode
          this._state.isEditingContact = false
          this._state.editedContact = null
          this.saveState()
        }
      }

      const email = {
        itemId: item.itemId,
        internetMessageId: null,
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
        // Extract Internet Message ID for reliable server-side lookups
        email.internetMessageId = email.internetHeaders['Message-ID'] || null
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

      this._currentItemId = email.itemId
      this._state.capturedEmail = email
      this.saveState()
      this.event('EmailCaptured', { subject: email.subject, itemId: email.itemId })

      // Get contact information from API
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
          try {
            field.getAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve(this._formatRecipients(result.value))
              } else {
                console.warn(fieldName + '.getAsync failed:', result.error)
                resolve([])
              }
            })
          } catch (e) {
            console.warn(fieldName + '.getAsync exception:', e)
            resolve([])
          }
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
          try {
            from.getAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                const v = result.value
                resolve(v ? { name: v.displayName || '', email: v.emailAddress || '' } : null)
              } else {
                console.warn('from.getAsync failed:', result.error)
                resolve(null)
              }
            })
          } catch (e) {
            console.warn('from.getAsync exception:', e)
            resolve(null)
          }
        })
      }
      return this._formatFrom(from)
    },
    async _getSubjectField(item) {
      const subject = item.subject
      if (subject && typeof subject === 'object' && subject.getAsync) {
        return new Promise((resolve) => {
          try {
            subject.getAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || '')
              } else {
                console.warn('subject.getAsync failed:', result.error)
                resolve('')
              }
            })
          } catch (e) {
            console.warn('subject.getAsync exception:', e)
            resolve('')
          }
        })
      }
      return typeof subject === 'string' ? subject : ''
    },
    async _getCategoriesField(item) {
      if (!item.categories) return []
      if (typeof item.categories === 'object' && item.categories.getAsync) {
        return new Promise((resolve) => {
          try {
            item.categories.getAsync((result) => {
              if (result.status === this.Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || [])
              } else {
                console.warn('categories.getAsync failed:', result.error)
                resolve([])
              }
            })
          } catch (e) {
            console.warn('categories.getAsync exception:', e)
            resolve([])
          }
        })
      }
      if (Array.isArray(item.categories)) return item.categories
      return []
    },
    async _getInternetHeaders(item) {
      if (!item.getAllInternetHeadersAsync) return {}
      return new Promise((resolve) => {
        try {
          item.getAllInternetHeadersAsync((result) => {
            if (result.status === this.Office.AsyncResultStatus.Succeeded && result.value) {
              resolve(this._parseHeaders(result.value))
            } else {
              console.warn('getAllInternetHeadersAsync failed:', result.error)
              resolve({})
            }
          })
        } catch (e) {
          console.warn('getAllInternetHeadersAsync exception:', e)
          resolve({})
        }
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
    _hasContactEdits() {
      // Check if there are any edits to the contact
      if (!this._state.isEditingContact || !this._state.editedContact || !this._state.contactInfo) {
        return false
      }

      const original = this._state.contactInfo
      const edited = this._state.editedContact

      // Compare all editable fields
      const fieldsToCompare = [
        'Firstname', 'Surname', 'JobTitle', 'Company', 'Mobile', 'AccountNo', 'EmailAddress',
        'Street1', 'Street2', 'City', 'PostCode', 'EmailNameAlias',
        'Linkedin', 'X', 'Facebook', 'Instagram', 'OtherChan1', 'OtherChan2',
        'SubscriberAttr1', 'SubscriberAttr2', 'SubscriberAttr3', 'SubscriberAttr4',
        'VIPStatus'
      ]

      for (const field of fieldsToCompare) {
        const origValue = original[field] || ''
        const editValue = edited[field] || ''

        // For VIPStatus, compare as booleans
        if (field === 'VIPStatus') {
          if (Boolean(origValue) !== Boolean(editValue)) {
            return true
          }
        } else {
          // For other fields, compare as strings
          if (String(origValue) !== String(editValue)) {
            return true
          }
        }
      }

      return false
    },
    async _showConfirmDialog(title, message, confirmText, cancelText) {
      // Create a custom confirmation dialog that works in Office Add-ins
      return new Promise((resolve) => {
        // Create overlay
        const overlay = document.createElement('div')
        overlay.className = 'confirm-overlay'

        // Create dialog
        const dialog = document.createElement('div')
        dialog.className = 'confirm-dialog'

        // Title
        const titleEl = document.createElement('div')
        titleEl.className = 'confirm-title'
        titleEl.textContent = title

        // Message
        const messageEl = document.createElement('div')
        messageEl.className = 'confirm-message'
        messageEl.textContent = message

        // Buttons container
        const buttonsEl = document.createElement('div')
        buttonsEl.className = 'confirm-buttons'

        // Confirm button
        const confirmBtn = document.createElement('button')
        confirmBtn.className = 'confirm-btn-ok'
        confirmBtn.textContent = confirmText || 'OK'
        confirmBtn.onclick = () => {
          document.body.removeChild(overlay)
          resolve(true)
        }

        // Cancel button
        const cancelBtn = document.createElement('button')
        cancelBtn.className = 'confirm-btn-cancel'
        cancelBtn.textContent = cancelText || 'Cancel'
        cancelBtn.onclick = () => {
          document.body.removeChild(overlay)
          resolve(false)
        }

        // Assemble
        buttonsEl.appendChild(cancelBtn)
        buttonsEl.appendChild(confirmBtn)
        dialog.appendChild(titleEl)
        dialog.appendChild(messageEl)
        dialog.appendChild(buttonsEl)
        overlay.appendChild(dialog)
        document.body.appendChild(overlay)

        // Focus confirm button
        confirmBtn.focus()
      })
    },
    async _showAlert(message) {
      // Create a custom alert dialog that works in Office Add-ins
      return new Promise((resolve) => {
        // Create overlay
        const overlay = document.createElement('div')
        overlay.className = 'confirm-overlay'

        // Create dialog
        const dialog = document.createElement('div')
        dialog.className = 'confirm-dialog'

        // Title
        const titleEl = document.createElement('div')
        titleEl.className = 'confirm-title'
        titleEl.textContent = 'Notice'

        // Message
        const messageEl = document.createElement('div')
        messageEl.className = 'confirm-message'
        messageEl.textContent = message

        // Buttons container
        const buttonsEl = document.createElement('div')
        buttonsEl.className = 'confirm-buttons'

        // OK button
        const okBtn = document.createElement('button')
        okBtn.className = 'confirm-btn-ok'
        okBtn.textContent = 'OK'
        okBtn.onclick = () => {
          document.body.removeChild(overlay)
          resolve()
        }

        // Assemble
        buttonsEl.appendChild(okBtn)
        dialog.appendChild(titleEl)
        dialog.appendChild(messageEl)
        dialog.appendChild(buttonsEl)
        overlay.appendChild(dialog)
        document.body.appendChild(overlay)

        // Focus OK button
        okBtn.focus()
      })
    },
    async _handleEditButtonClick() {
      // Handle edit button click when already in edit mode
      if (!this._state.isEditingContact) {
        // Not in edit mode, start editing
        this.startEditingContact()
        return
      }

      // Already in edit mode - check for edits
      const hasEdits = this._hasContactEdits()

      if (hasEdits) {
        // Has edits, prompt user
        const shouldSave = await this._showConfirmDialog(
          'Unsaved Changes',
          'You have unsaved contact changes. Do you want to save them?',
          'Save',
          'Discard'
        )

        if (shouldSave) {
          // Save changes and exit edit mode
          await this.saveEditedContact()
        } else {
          // Discard changes and exit edit mode
          this.cancelEditingContact()
        }
      } else {
        // No edits, silently exit edit mode
        this.cancelEditingContact()
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
    _createContactFieldEdit(fieldName, label, value, disabled) {
      const val = value || ''
      const dis = disabled ? ' disabled' : ''

      let html = '<div class="contact-field-edit">'
      html += '<span class="field-label">' + this._escapeHtml(label) + ':</span>'
      html += '<input type="text" id="edit_' + fieldName + '" value="' + this._escapeHtml(val) + '"' + dis + '>'
      html += '</div>'

      return html
    },
    _createContactFieldAlwaysShow(label, value, isUrl) {
      let html = '<div class="contact-field">'
      html += '<span class="field-label">' + this._escapeHtml(label) + ':</span> '

      if (value) {
        if (isUrl && this._isUrl(value)) {
          html += '<a href="' + this._escapeHtml(value) + '" target="_blank" rel="noopener noreferrer">' +
            this._escapeHtml(value) + '</a>'
        } else {
          html += '<span class="field-value">' + this._escapeHtml(value) + '</span>'
        }
      } else {
        html += '<span class="field-value field-empty">-</span>'
      }

      html += '</div>'

      return html
    },
    _formatRecipientsList(recipients) {
      if (!recipients || !Array.isArray(recipients) || recipients.length === 0) {
        return '-'
      }
      return recipients.map(r => {
        const name = r.name || ''
        const email = r.email || ''
        if (name && email) {
          return name + ' <' + email + '>'
        }
        return email || name || 'Unknown'
      }).join(', ')
    },
    _attachContactEventListeners() {
      // Attach edit button listener
      const editBtn = document.getElementById('editContactBtn')
      if (editBtn) {
        editBtn.onclick = () => {
          this._handleEditButtonClick()
        }
      }

      // Attach chevron button listener
      const chevronBtn = document.getElementById('toggleChevronBtn')
      if (chevronBtn) {
        chevronBtn.onclick = () => {
          this.toggleMoreContact()
        }
      }
    },
    _updateUI() {
      if (typeof document === 'undefined') return

      const userNameEl = document.getElementById('userName')
      const folderNameEl = document.getElementById('folderName')
      const platformEl = document.getElementById('platform')
      const versionEl = document.getElementById('version')
      const contactSectionEl = document.getElementById('contactSection')
      const conversationSummarySectionEl = document.getElementById('conversationSummarySection')
      const actionsSectionEl = document.getElementById('actionsSection')
      const sectionTitleEls = document.querySelectorAll('.section-title-container')

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

      // Update Contact section title with edit and chevron buttons
      if (sectionTitleEls.length > 0) {
        const contactTitleEl = sectionTitleEls[0]
        const contact = this._state.contactInfo
        if (contact) {
          const chevronClass = this._state.showMoreContact ? 'chevron-up' : 'chevron-down'
          contactTitleEl.innerHTML = '<span class="section-title">Contact</span>' +
            '<div class="section-title-buttons">' +
            '<button id="editContactBtn" class="edit-btn" title="Edit contact">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
            '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>' +
            '<path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>' +
            '</svg>' +
            '</button>' +
            '<button id="toggleChevronBtn" class="chevron-btn ' + chevronClass + '" title="Toggle contact details">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
            '<polyline points="6 9 12 15 18 9"></polyline>' +
            '</svg>' +
            '</button>' +
            '</div>'
        } else {
          contactTitleEl.innerHTML = '<span class="section-title">Contact</span>'
        }
      }

      // Update contact section
      if (contactSectionEl) {
        const contact = this._state.contactInfo
        const isEditing = this._state.isEditingContact
        const editedContact = this._state.editedContact

        if (contact) {
          if (isEditing && editedContact) {
            // Edit mode
            let html = '<div class="contact-card">'

            // Name fields with labels - vertical layout
            html += '<div class="contact-name-edit">'
            html += '<div class="contact-name-field">'
            html += '<span class="field-label">First Name:</span>'
            html += '<input type="text" id="edit_Firstname" placeholder="First Name" value="' +
              this._escapeHtml(editedContact.Firstname || '') + '">'
            html += '</div>'
            html += '<div class="contact-name-field">'
            html += '<span class="field-label">Surname:</span>'
            html += '<input type="text" id="edit_Surname" placeholder="Surname" value="' +
              this._escapeHtml(editedContact.Surname || '') + '">'
            html += '</div>'
            html += '</div>'

            // VIP Status checkbox
            html += '<div class="contact-field-edit vip-field">'
            html += '<label class="vip-checkbox-label">'
            html += '<input type="checkbox" id="edit_VIPStatus"' + (editedContact.VIPStatus ? ' checked' : '') + '>'
            html += '<span class="vip-checkbox-text">&nbsp;&nbsp;VIP Status</span>'
            html += '</label>'
            html += '</div>'

            html += this._createContactFieldEdit('JobTitle', 'Job Title', editedContact.JobTitle, false)
            html += this._createContactFieldEdit('Company', 'Company', editedContact.Company, false)
            html += this._createContactFieldEdit('Mobile', 'Mobile', editedContact.Mobile, false)
            html += this._createContactFieldEdit('AccountNo', 'Account No', editedContact.AccountNo, false)
            html += this._createContactFieldEdit('EmailAddress', 'Email', editedContact.EmailAddress, false)

            // Always show all fields in edit mode
            html += '<div class="contact-more">'
            html += this._createContactFieldEdit('UID', 'UID', editedContact.UID, true)
            html += this._createContactFieldEdit('Street1', 'Street1', editedContact.Street1, false)
            html += this._createContactFieldEdit('Street2', 'Street2', editedContact.Street2, false)
            html += this._createContactFieldEdit('City', 'City', editedContact.City, false)
            html += this._createContactFieldEdit('PostCode', 'PostCode', editedContact.PostCode, false)
            html += this._createContactFieldEdit('EmailNameAlias', 'Email Alias', editedContact.EmailNameAlias, false)
            html += this._createContactFieldEdit('Linkedin', 'LinkedIn', editedContact.Linkedin, false)
            html += this._createContactFieldEdit('X', 'X', editedContact.X, false)
            html += this._createContactFieldEdit('Facebook', 'Facebook', editedContact.Facebook, false)
            html += this._createContactFieldEdit('Instagram', 'Instagram', editedContact.Instagram, false)
            html += this._createContactFieldEdit('OtherChan1', 'Other Channel 1', editedContact.OtherChan1, false)
            html += this._createContactFieldEdit('OtherChan2', 'Other Channel 2', editedContact.OtherChan2, false)
            html += this._createContactFieldEdit('SubscriberAttr1', 'Subscriber Attr 1', editedContact.SubscriberAttr1, false)
            html += this._createContactFieldEdit('SubscriberAttr2', 'Subscriber Attr 2', editedContact.SubscriberAttr2, false)
            html += this._createContactFieldEdit('SubscriberAttr3', 'Subscriber Attr 3', editedContact.SubscriberAttr3, false)
            html += this._createContactFieldEdit('SubscriberAttr4', 'Subscriber Attr 4', editedContact.SubscriberAttr4, false)

            // Timestamp fields (read-only)
            if (editedContact.CreatedAt) {
              html += this._createContactField('Created', this._formatTimestamp(editedContact.CreatedAt), false)
            }
            if (editedContact.UpdatedAt) {
              html += this._createContactField('Updated', this._formatTimestamp(editedContact.UpdatedAt), false)
            }
            if (editedContact.LastContactedAt) {
              html += this._createContactField('Last Contacted', this._formatTimestamp(editedContact.LastContactedAt), false)
            }
            html += '</div>'

            html += '<div class="contact-edit-buttons">'
            html += '<button id="saveContactBtn" class="save-btn">Save Changes</button>'
            html += '<button id="cancelContactBtn" class="cancel-btn">Cancel</button>'
            html += '</div>'
            html += '</div>'

            contactSectionEl.innerHTML = html

            // Attach VIP checkbox listener
            const vipCheckbox = document.getElementById('edit_VIPStatus')
            if (vipCheckbox) {
              vipCheckbox.addEventListener('change', (e) => {
                this.updateEditedContactVIPStatus(e.target.checked)
              })
            }

            // Attach input change listeners - now includes EmailAddress
            const editableFields = [
              'Firstname', 'Surname', 'JobTitle', 'Company', 'Mobile', 'AccountNo', 'EmailAddress',
              'Street1', 'Street2', 'City', 'PostCode', 'EmailNameAlias',
              'Linkedin', 'X', 'Facebook', 'Instagram', 'OtherChan1', 'OtherChan2',
              'SubscriberAttr1', 'SubscriberAttr2', 'SubscriberAttr3', 'SubscriberAttr4'
            ]

            editableFields.forEach((fieldName) => {
              const input = document.getElementById('edit_' + fieldName)
              if (input) {
                input.addEventListener('input', (e) => {
                  this.updateEditedContactField(fieldName, e.target.value)
                })
              }
            })

            // Attach save button listener
            const saveBtn = document.getElementById('saveContactBtn')
            if (saveBtn) {
              saveBtn.onclick = () => {
                this.saveEditedContact()
              }
            }

            // Attach cancel button listener
            const cancelBtn = document.getElementById('cancelContactBtn')
            if (cancelBtn) {
              cancelBtn.onclick = () => {
                this.cancelEditingContact()
              }
            }
          } else {
            // View mode
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
            html += this._createContactFieldAlwaysShow('Email', contact.EmailAddress, false)

            // More/Less toggle
            if (this._state.showMoreContact) {
              html += '<div class="contact-more">'
              html += this._createContactField('UID', contact.UID, false)
              html += this._createContactField('Street1', contact.Street1, false)
              html += this._createContactField('Street2', contact.Street2, false)
              html += this._createContactField('City', contact.City, false)
              html += this._createContactField('PostCode', contact.PostCode, false)
              html += this._createContactField('Email Alias', contact.EmailNameAlias, false)

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
            }

            html += '</div>'
            contactSectionEl.innerHTML = html
          }

          // FIXED: Always attach event listeners after updating innerHTML
          this._attachContactEventListeners()
        } else {
          contactSectionEl.innerHTML = '<div class="no-contact">No contact information available</div>'
        }
      }

      // Update conversation summary section
      if (conversationSummarySectionEl) {
        conversationSummarySectionEl.innerHTML = '<div class="no-conversation-summary">No conversation summary available</div>'
      }

      // Update email attributes section
      this._updateEmailAttributesUI()

      // Update actions section
      if (actionsSectionEl) {
        actionsSectionEl.innerHTML = '<div class="no-actions">No actions available</div>'
      }
    },
    _updateEmailAttributesUI() {
      if (typeof document === 'undefined') return

      const emailAttributesSectionEl = document.getElementById('emailAttributesSection')
      if (!emailAttributesSectionEl) return

      const email = this._state.capturedEmail

      if (email) {
        let html = '<div class="email-attributes-container">'

        // Importance
        html += '<div class="email-attribute">'
        html += '<span class="field-label">Importance:</span> '
        const importance = email.importance || 'normal'
        let importanceClass = 'importance-normal'
        if (importance === 'high') importanceClass = 'importance-high'
        else if (importance === 'low') importanceClass = 'importance-low'
        html += '<span class="importance-badge ' + importanceClass + '">' + this._escapeHtml(importance) + '</span>'
        html += '</div>'

        // Sentiment
        html += '<div class="email-attribute">'
        html += '<span class="field-label">Sentiment:</span> '
        const sentiment = email.sentiment || 'neutral'
        let sentimentClass = 'sentiment-neutral'
        if (sentiment === 'positive') sentimentClass = 'sentiment-positive'
        else if (sentiment === 'negative') sentimentClass = 'sentiment-negative'
        html += '<span class="sentiment-badge ' + sentimentClass + '">' + this._escapeHtml(sentiment) + '</span>'
        html += '</div>'

        // Categories
        if (email.categories && email.categories.length > 0) {
          html += '<div class="email-attribute">'
          html += '<span class="field-label">Categories:</span> '
          html += '<div class="category-tags">'
          email.categories.forEach(function(cat) {
            html += '<span class="category-tag">' + this._escapeHtml(cat) + '</span>'
          }.bind(this))
          html += '</div>'
          html += '</div>'
        }

        html += '</div>'
        emailAttributesSectionEl.innerHTML = html
      } else {
        emailAttributesSectionEl.innerHTML = '<div class="no-email-attributes">No email selected</div>'
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