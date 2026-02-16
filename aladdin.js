// aladdin.js

const singleton = [false]
const STORAGE_KEY = 'aladdin_state'
const CATEGORY_INTERVAL_MS = 8 * 60 * 60 * 1000 // 8 hours
const API_BASE = 'https://www.devappeggio.com'

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
    _userName: '',
    _folderName: 'Inbox',
    _platform: '',
    _version: '',
    _itemHandlersRegistered: false,
    _state: {
      events: [],
      capturedEmail: null,
      _lastCategoryInit: null,
    },
    state() {
      return this._state
    },
    saveState() {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(this._state))
      } catch (e) {
        console.error('Failed to save state', e)
      }
    },
    loadState() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY)
        if (raw) {
          const parsed = JSON.parse(raw)
          this._state = parsed
          if (!this._state.events) this._state.events = []
          if (!this._state.capturedEmail) this._state.capturedEmail = null
        }
      } catch (e) {
        console.error('Failed to load state', e)
      }
    },
    watchState() {
      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (e) => {
          if (e.key === STORAGE_KEY && e.newValue) {
            try {
              const parsed = JSON.parse(e.newValue)
              this._state = parsed
              if (!this._state.events) this._state.events = []
              this.updateUI()
            } catch (err) {
              console.error('Failed to parse storage event', err)
            }
          }
        })
      }
    },
    event(name, details) {
      console.log('Event:', name)
      const entry = {
        name: name,
        timestamp: new Date().toISOString(),
        details: details || null,
      }
      this._state.events.push(entry)
      if (this._state.events.length > 10) {
        this._state.events = this._state.events.slice(-10)
      }
      this.saveState()
      this.updateUI()
    },
    initialize() {
      this._userName = this._getUserName()
      this._platform = this._getPlatform()
      this._version = this._getVersion()

      this.event('addin-initialized', {
        user: this._userName,
        platform: this._platform,
        version: this._version,
      })

      // Check for previous email that needs notification (Rule R3)
      this._handleDeselectionOnInit()

      // Register mailbox-level events
      this._registerMailboxEvents()

      // Process current item
      this._processCurrentItem()

      // Init categories (Rule R4)
      this._initCategories()

      this.updateUI()
    },
    _getUserName() {
      try {
        return this.Office.context.mailbox.userProfile.displayName || 'Unknown User'
      } catch (e) {
        return 'Unknown User'
      }
    },
    _getPlatform() {
      try {
        return String(this.Office.context.platform || 'Unknown')
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
    _handleDeselectionOnInit() {
      // Rule R3: re-read state from localStorage
      this.loadState()
      const saved = this._state.capturedEmail
      if (!saved) return

      const item = this.Office.context.mailbox.item
      const isCompose = item && !item.itemId
      const noItem = !item
      const differentItem = item && item.itemId && saved.graphMessageId &&
        this._getItemIdForComparison(item) !== saved.graphMessageId

      if (isCompose || noItem || differentItem) {
        // Rule R2: re-read state before notify
        this.loadState()
        if (this._state.capturedEmail) {
          this.notify(this._state.capturedEmail)
          this._state.capturedEmail = null
          this.saveState()
        }
      }
    },
    _getItemIdForComparison(item) {
      if (!item || !item.itemId) return null
      try {
        if (this.Office.context.mailbox.convertToRestId) {
          return this.Office.context.mailbox.convertToRestId(
            item.itemId,
            this.Office.MailboxEnums.RestVersion.v2_0
          )
        }
      } catch (e) { }
      return item.itemId
    },
    _registerMailboxEvents() {
      const mailbox = this.Office.context.mailbox
      // Mailbox-level: ItemChanged
      try {
        mailbox.addHandlerAsync(
          this.Office.EventType.ItemChanged,
          (eventArgs) => this._onItemChanged(eventArgs)
        )
      } catch (e) {
        console.error('Failed to register ItemChanged', e)
      }
    },
    _registerItemEvents() {
      const item = this.Office.context.mailbox.item
      if (!item) return
      if (this._itemHandlersRegistered) return

      const itemEvents = [
        'RecipientsChanged',
        'AttachmentsChanged',
        'RecurrenceChanged',
        'AppointmentTimeChanged',
        'EnhancedLocationsChanged',
      ]

      itemEvents.forEach((evtName) => {
        try {
          if (this.Office.EventType[evtName]) {
            item.addHandlerAsync(
              this.Office.EventType[evtName],
              (eventArgs) => {
                this.event(evtName, eventArgs || {})
                // Re-capture on relevant changes
                if (evtName === 'RecipientsChanged' || evtName === 'AttachmentsChanged') {
                  this._captureCurrentEmail()
                }
              }
            )
          }
        } catch (e) {
          console.error('Failed to register ' + evtName, e)
        }
      })

      this._itemHandlersRegistered = true
    },
    _onItemChanged(eventArgs) {
      this.event('ItemChanged', eventArgs || {})
      this._itemHandlersRegistered = false

      // Rule R2: re-read state before notifying
      this.loadState()
      const previousEmail = this._state.capturedEmail

      const item = this.Office.context.mailbox.item
      const isCompose = item && !item.itemId

      // Rule R1: compose is explicit deselection
      if (isCompose) {
        if (previousEmail) {
          this.notify(previousEmail)
          this._state.capturedEmail = null
          this._currentItemId = null
          this.saveState()
        }
        this._extractFolderFromHeaders()
        this._registerItemEvents()
        this.updateUI()
        return
      }

      if (!item) {
        // No item selected
        if (previousEmail) {
          this.notify(previousEmail)
          this._state.capturedEmail = null
          this._currentItemId = null
          this.saveState()
        }
        this.updateUI()
        return
      }

      // Item exists with itemId - different email
      const newId = this._getItemIdForComparison(item)
      if (previousEmail && newId !== previousEmail.graphMessageId) {
        this.notify(previousEmail)
        this._state.capturedEmail = null
        this.saveState()
      }

      this._currentItemId = newId
      this._registerItemEvents()
      this._captureCurrentEmail()
      this._extractFolderFromHeaders()
    },
    _processCurrentItem() {
      const item = this.Office.context.mailbox.item
      if (!item) return

      const isCompose = !item.itemId

      // Rule R1: compose is explicit deselection (handled in _handleDeselectionOnInit)
      if (isCompose) {
        this._extractFolderFromHeaders()
        this._registerItemEvents()
        return
      }

      this._currentItemId = this._getItemIdForComparison(item)
      this._registerItemEvents()
      this._captureCurrentEmail()
      this._extractFolderFromHeaders()
    },
    _captureCurrentEmail() {
      const item = this.Office.context.mailbox.item
      if (!item || !item.itemId) return

      const emailData = {
        graphMessageId: this._getItemIdForComparison(item),
        to: [],
        from: null,
        cc: [],
        subject: '',
        attachments: [],
        categories: [],
        importance: '',
        sentiment: '',
        internetHeaders: {},
        capturedAt: new Date().toISOString(),
      }

      let pending = 0
      const self = this

      function done() {
        pending--
        if (pending <= 0) {
          self._state.capturedEmail = emailData
          self.saveState()
          self.updateUI()
        }
      }

      // To
      pending++
      if (item.to && typeof item.to.getAsync === 'function') {
        item.to.getAsync((result) => {
          if (result.status === self.Office.AsyncResultStatus.Succeeded) {
            emailData.to = (result.value || []).map(r => ({ name: r.displayName, email: r.emailAddress }))
          }
          done()
        })
      } else if (item.to) {
        emailData.to = (Array.isArray(item.to) ? item.to : []).map(r => ({ name: r.displayName, email: r.emailAddress }))
        done()
      } else {
        done()
      }

      // From
      pending++
      if (item.from && typeof item.from.getAsync === 'function') {
        item.from.getAsync((result) => {
          if (result.status === self.Office.AsyncResultStatus.Succeeded && result.value) {
            emailData.from = { name: result.value.displayName, email: result.value.emailAddress }
          }
          done()
        })
      } else if (item.from) {
        emailData.from = { name: item.from.displayName, email: item.from.emailAddress }
        done()
      } else if (item.sender) {
        emailData.from = { name: item.sender.displayName, email: item.sender.emailAddress }
        done()
      } else {
        done()
      }

      // CC
      pending++
      if (item.cc && typeof item.cc.getAsync === 'function') {
        item.cc.getAsync((result) => {
          if (result.status === self.Office.AsyncResultStatus.Succeeded) {
            emailData.cc = (result.value || []).map(r => ({ name: r.displayName, email: r.emailAddress }))
          }
          done()
        })
      } else if (item.cc) {
        emailData.cc = (Array.isArray(item.cc) ? item.cc : []).map(r => ({ name: r.displayName, email: r.emailAddress }))
        done()
      } else {
        done()
      }

      // Subject
      pending++
      if (item.subject && typeof item.subject.getAsync === 'function') {
        item.subject.getAsync((result) => {
          if (result.status === self.Office.AsyncResultStatus.Succeeded) {
            emailData.subject = result.value || ''
          }
          done()
        })
      } else if (typeof item.subject === 'string') {
        emailData.subject = item.subject
        done()
      } else {
        done()
      }

      // Attachments
      pending++
      if (item.attachments) {
        emailData.attachments = (Array.isArray(item.attachments) ? item.attachments : []).map(a => a.name)
      }
      done()

      // Categories
      pending++
      if (item.categories && typeof item.categories.getAsync === 'function') {
        item.categories.getAsync((result) => {
          if (result.status === self.Office.AsyncResultStatus.Succeeded) {
            emailData.categories = (result.value || []).map(c => c.displayName || c)
          }
          done()
        })
      } else if (item.categories) {
        emailData.categories = Array.isArray(item.categories) ? item.categories : []
        done()
      } else {
        done()
      }

      // Importance
      pending++
      if (typeof item.importance === 'string') {
        emailData.importance = item.importance
        done()
      } else {
        emailData.importance = 'normal'
        done()
      }

      // Internet Headers (for sentiment and folder)
      pending++
      if (item.getAllInternetHeadersAsync) {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === self.Office.AsyncResultStatus.Succeeded && result.value) {
            emailData.internetHeaders = self._parseInternetHeaders(result.value)
            emailData.sentiment = self._extractSentiment(emailData.internetHeaders)
            const folderFromHeaders = self._extractFolderFromParsedHeaders(emailData.internetHeaders)
            if (folderFromHeaders) {
              self._folderName = folderFromHeaders
            }
          }
          done()
        })
      } else {
        done()
      }
    },
    _parseInternetHeaders(headerString) {
      const headers = {}
      if (!headerString) return headers
      // Unfold headers (lines starting with whitespace are continuations)
      const unfolded = headerString.replace(/\r?\n[ \t]+/g, ' ')
      const lines = unfolded.split(/\r?\n/)
      lines.forEach((line) => {
        const colonIndex = line.indexOf(':')
        if (colonIndex > 0) {
          const key = line.substring(0, colonIndex).trim()
          const value = line.substring(colonIndex + 1).trim()
          if (headers[key]) {
            if (Array.isArray(headers[key])) {
              headers[key].push(value)
            } else {
              headers[key] = [headers[key], value]
            }
          } else {
            headers[key] = value
          }
        }
      })
      return headers
    },
    _extractSentiment(headers) {
      // Check common sentiment-related headers
      const sentimentHeaders = [
        'X-MS-Exchange-Organization-SCL',
        'X-Spam-Status',
        'X-Spam-Score',
        'X-Sentiment',
        'X-MS-Exchange-Organization-AuthAs',
      ]
      for (let i = 0; i < sentimentHeaders.length; i++) {
        if (headers[sentimentHeaders[i]]) {
          return String(headers[sentimentHeaders[i]])
        }
      }
      return ''
    },
    _extractFolderFromParsedHeaders(headers) {
      // Try various headers that might contain folder info
      const folderHeaders = [
        'X-Folder-Name',
        'X-Folder',
        'X-MS-Exchange-Organization-FolderName',
        'X-GM-LABELS',
      ]
      for (let i = 0; i < folderHeaders.length; i++) {
        if (headers[folderHeaders[i]]) {
          return String(headers[folderHeaders[i]])
        }
      }
      // Try to extract from Delivered-To or other routing headers
      if (headers['X-Microsoft-Antispam-Mailbox-Delivery']) {
        const delivery = String(headers['X-Microsoft-Antispam-Mailbox-Delivery'])
        const folderMatch = delivery.match(/dest:([^;]+)/)
        if (folderMatch) return folderMatch[1].trim()
      }
      return null
    },
    _extractFolderFromHeaders() {
      const item = this.Office.context.mailbox.item
      if (!item) return
      if (item.getAllInternetHeadersAsync) {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === this.Office.AsyncResultStatus.Succeeded && result.value) {
            const headers = this._parseInternetHeaders(result.value)
            const folder = this._extractFolderFromParsedHeaders(headers)
            if (folder) {
              this._folderName = folder
              this.updateUI()
            }
          }
        })
      }
    },
    async notify(emailData) {
      if (!emailData) return
      this.event('notify-sent', { graphMessageId: emailData.graphMessageId, subject: emailData.subject })
      try {
        await fetch(API_BASE + '/api/inboxnotify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            user: this._userName,
            folder: this._folderName,
            platform: this._platform,
            version: this._version,
            email: emailData,
          }),
        })
      } catch (e) {
        console.error('Failed to notify API', e)
      }
    },
    async getCategories(initInfo) {
      try {
        const response = await fetch(API_BASE + '/api/inboxinit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(initInfo),
        })
        if (response.ok) {
          const data = await response.json()
          if (Array.isArray(data)) return data
        }
      } catch (e) {
        console.error('Failed to fetch categories from API', e)
      }
      // Default fallback
      return [
        { displayName: 'InboxAgent', color: 'Preset0' }
      ]
    },
    async _initCategories() {
      // Rule R4: only run once every 8 hours
      this.loadState()
      const lastInit = this._state._lastCategoryInit
      if (lastInit && (Date.now() - lastInit) < CATEGORY_INTERVAL_MS) {
        return
      }

      const initInfo = {
        user: this._userName,
        folder: this._folderName,
        platform: this._platform,
        version: this._version,
      }

      const categories = await this.getCategories(initInfo)
      this._ensureMasterCategories(categories)

      this._state._lastCategoryInit = Date.now()
      this.saveState()
    },
    _ensureMasterCategories(categories) {
      if (!categories || !categories.length) return
      const mailbox = this.Office.context.mailbox
      if (!mailbox.masterCategories || !mailbox.masterCategories.addAsync) {
        console.error('masterCategories.addAsync not available')
        return
      }

      // Map categories to the format expected by the API
      const masterCategories = categories.map((cat) => ({
        displayName: cat.displayName,
        color: cat.color || 'Preset0',
      }))

      mailbox.masterCategories.addAsync(masterCategories, (result) => {
        if (result.status === this.Office.AsyncResultStatus.Succeeded) {
          this.event('categories-added', { categories: masterCategories.map(c => c.displayName) })
        } else {
          // May fail if categories already exist, which is fine
          console.log('masterCategories.addAsync result:', result.status, result.error)
        }
      })
    },
    updateUI() {
      if (typeof document === 'undefined') return

      // User info section
      const userInfoEl = document.getElementById('user-info')
      if (userInfoEl) {
        userInfoEl.innerHTML = ''
        const nameEl = document.createElement('div')
        nameEl.className = 'user-name'
        nameEl.textContent = this._userName || 'Unknown User'
        userInfoEl.appendChild(nameEl)

        const folderEl = document.createElement('div')
        folderEl.className = 'user-detail'
        folderEl.textContent = 'Folder: ' + (this._folderName || 'Inbox')
        userInfoEl.appendChild(folderEl)

        const platformEl = document.createElement('div')
        platformEl.className = 'user-detail'
        platformEl.textContent = 'Platform: ' + (this._platform || 'Unknown')
        userInfoEl.appendChild(platformEl)

        const versionEl = document.createElement('div')
        versionEl.className = 'user-detail'
        versionEl.textContent = 'Version: ' + (this._version || 'Unknown')
        userInfoEl.appendChild(versionEl)
      }

      // Current email section
      const emailInfoEl = document.getElementById('email-info')
      if (emailInfoEl) {
        const email = this._state.capturedEmail
        if (email) {
          emailInfoEl.innerHTML = ''
          const subjectEl = document.createElement('div')
          subjectEl.className = 'email-subject'
          subjectEl.textContent = email.subject || '(No Subject)'
          emailInfoEl.appendChild(subjectEl)

          const fromEl = document.createElement('div')
          fromEl.className = 'email-detail'
          fromEl.textContent = 'From: ' + (email.from ? (email.from.name || email.from.email) : 'Unknown')
          emailInfoEl.appendChild(fromEl)

          const toEl = document.createElement('div')
          toEl.className = 'email-detail'
          toEl.textContent = 'To: ' + (email.to || []).map(r => r.name || r.email).join(', ')
          emailInfoEl.appendChild(toEl)

          if (email.cc && email.cc.length > 0) {
            const ccEl = document.createElement('div')
            ccEl.className = 'email-detail'
            ccEl.textContent = 'CC: ' + email.cc.map(r => r.name || r.email).join(', ')
            emailInfoEl.appendChild(ccEl)
          }

          if (email.attachments && email.attachments.length > 0) {
            const attEl = document.createElement('div')
            attEl.className = 'email-detail'
            attEl.textContent = 'Attachments: ' + email.attachments.join(', ')
            emailInfoEl.appendChild(attEl)
          }

          if (email.categories && email.categories.length > 0) {
            const catEl = document.createElement('div')
            catEl.className = 'email-detail'
            catEl.textContent = 'Categories: ' + email.categories.join(', ')
            emailInfoEl.appendChild(catEl)
          }

          const impEl = document.createElement('div')
          impEl.className = 'email-detail'
          impEl.textContent = 'Importance: ' + (email.importance || 'normal')
          emailInfoEl.appendChild(impEl)

          if (email.sentiment) {
            const sentEl = document.createElement('div')
            sentEl.className = 'email-detail'
            sentEl.textContent = 'Sentiment: ' + email.sentiment
            emailInfoEl.appendChild(sentEl)
          }

          const idEl = document.createElement('div')
          idEl.className = 'email-detail email-id'
          idEl.textContent = 'ID: ' + (email.graphMessageId || 'N/A')
          emailInfoEl.appendChild(idEl)
        } else {
          emailInfoEl.innerHTML = '<div class="no-email">No email selected</div>'
        }
      }

      // Events section
      const eventsEl = document.getElementById('events')
      if (eventsEl) {
        if (!this._state.events || this._state.events.length === 0) {
          eventsEl.innerHTML = '<div class="no-events">No events recorded yet</div>'
        } else {
          eventsEl.innerHTML = ''
          const reversed = [...this._state.events].reverse()
          reversed.forEach((evt) => {
            const itemEl = document.createElement('div')
            itemEl.className = 'event-item'

            const nameEl = document.createElement('div')
            nameEl.className = 'event-name'
            nameEl.textContent = evt.name
            itemEl.appendChild(nameEl)

            const tsEl = document.createElement('div')
            tsEl.className = 'event-timestamp'
            tsEl.textContent = evt.timestamp
            itemEl.appendChild(tsEl)

            if (evt.details) {
              const detEl = document.createElement('div')
              detEl.className = 'event-details'
              detEl.textContent = JSON.stringify(evt.details, null, 2)
              itemEl.appendChild(detEl)
            }

            eventsEl.appendChild(itemEl)
          })
        }
      }

      // Status
      const statusEl = document.getElementById('status')
      if (statusEl) {
        statusEl.textContent = 'Ready'
        statusEl.className = 'status-ready'
      }
    },
  }
}