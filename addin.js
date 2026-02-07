// addin.js

const singleton = [false]
const STATE_KEY = 'aladdin-addin-state'

export function createAddIn(Office) {
  if (typeof window !== 'undefined' && window.aladdinInstance) {
    window.aladdinInstance.loadState()
    return window.aladdinInstance;
  }
  if (singleton[0]) {
    singleton[0].loadState()
    return singleton[0];
  }
  const queue = new Queue()
  const instance = addin(queue, Office)
  if (typeof window !== 'undefined') window.aladdinInstance = instance;
  if (typeof window === 'undefined') singleton[0] = instance;
  instance.loadState()
  return instance
}

function addin(queue, Office) {
  return {
    queue() {
      return queue
    },
    start() {
      queue.addEventListener('success', e => {
        console.log('Job ok:', JSON.stringify(e.detail, null, 2))
      })
      queue.addEventListener('error', e => {
        console.error('Job err:', e)
      })
      queue.start(err => {
        if (err) console.error(err)
      })
    },
    Office,
    _state: {
      globalData: {},
      eventCounts: {
        commands: 0,
        launchEvents: 0,
        itemChanges: 0
      }
    },
    state() {
      return this._state
    },
    saveState() {
      if (typeof localStorage !== 'undefined') {
        try {
          const stateJson = JSON.stringify(this._state)
          localStorage.setItem(STATE_KEY, stateJson)
          console.log('State saved to localStorage')
        } catch (error) {
          console.error('Error saving state to localStorage:', error)
        }
      } else {
        console.warn('localStorage not available')
      }
    },
    loadState() {
      if (typeof localStorage !== 'undefined') {
        try {
          const stateJson = localStorage.getItem(STATE_KEY)
          if (stateJson) {
            this._state = JSON.parse(stateJson)
            console.log('State loaded from localStorage:', this._state)
          } else {
            console.log('No saved state found in localStorage')
          }
        } catch (error) {
          console.error('Error loading state from localStorage:', error)
        }
      } else {
        console.warn('localStorage not available')
      }
    },
    changeState(changes) {
      if (changes.eventCounts) Object.assign(this._state.eventCounts, changes.eventCounts);
      if (changes.globalData) Object.assign(this._state.globalData, changes.globalData);
      this.saveState()
    }
  }
}

const has = Object.prototype.hasOwnProperty

export class QueueEvent extends Event {
  constructor (name, detail) {
    super(name)
    this.detail = detail
  }
}

// Asynchronous function queue with adjustable concurrency.
// a class Queue that implements most of the Array API. Pass async functions (ones that accept a callback or return a promise)
// to an instance's additive array methods. Processing begins when you call q.start().
export class Queue extends EventTarget {
  constructor (options = {}) {
    super()
    const { concurrency = Infinity, timeout = 0, autostart = false, results = null } = options
    this.concurrency = concurrency
    this.timeout = timeout
    this.autostart = autostart
    this.results = results
    this.pending = 0
    this.session = 0
    this.running = false
    this.jobs = []
    this.timers = []
    this.addEventListener('error', this._errorHandler)
  }

  _errorHandler (evt) {
    this.end(evt.detail.error)
  }

  pop () {
    return this.jobs.pop()
  }

  shift () {
    return this.jobs.shift()
  }

  indexOf (searchElement, fromIndex) {
    return this.jobs.indexOf(searchElement, fromIndex)
  }

  lastIndexOf (searchElement, fromIndex) {
    if (fromIndex !== undefined) return this.jobs.lastIndexOf(searchElement, fromIndex)
    return this.jobs.lastIndexOf(searchElement)
  }

  slice (start, end) {
    this.jobs = this.jobs.slice(start, end)
    return this
  }

  reverse () {
    this.jobs.reverse()
    return this
  }

  push (...workers) {
    const methodResult = this.jobs.push(...workers)
    if (this.autostart) this._start()
    return methodResult
  }

  unshift (...workers) {
    const methodResult = this.jobs.unshift(...workers)
    if (this.autostart) this._start()
    return methodResult
  }

  splice (start, deleteCount, ...workers) {
    this.jobs.splice(start, deleteCount, ...workers)
    if (this.autostart) this._start()
    return this
  }

  get length () {
    return this.pending + this.jobs.length
  }

  start (callback) {
    if (this.running) throw new Error('already started')
    let awaiter
    if (callback) {
      this._addCallbackToEndEvent(callback)
    } else {
      awaiter = this._createPromiseToEndEvent()
    }
    this._start()
    return awaiter
  }

  _start () {
    this.running = true
    if (this.pending >= this.concurrency) {
      return
    }
    if (this.jobs.length === 0) {
      if (this.pending === 0) {
        this.done()
      }
      return
    }
    const job = this.jobs.shift()
    const session = this.session
    const timeout = (job !== undefined) && has.call(job, 'timeout') ? job.timeout : this.timeout
    let once = true
    let timeoutId = null
    let didTimeout = false
    let resultIndex = null
    const next = (error, ...result) => {
      if (once && this.session === session) {
        once = false
        this.pending--
        if (timeoutId !== null) {
          this.timers = this.timers.filter(tID => tID !== timeoutId)
          clearTimeout(timeoutId)
        }
        if (error) {
          this.dispatchEvent(new QueueEvent('error', { error, job }))
        } else if (!didTimeout) {
          if (resultIndex !== null && this.results !== null) {
            this.results[resultIndex] = [...result]
          }
          this.dispatchEvent(new QueueEvent('success', { result: [...result], job }))
        }
        if (this.session === session) {
          if (this.pending === 0 && this.jobs.length === 0) {
            this.done()
          } else if (this.running) {
            this._start()
          }
        }
      }
    }
    if (timeout) {
      timeoutId = setTimeout(() => {
        didTimeout = true
        this.dispatchEvent(new QueueEvent('timeout', { next, job }))
        next()
      }, timeout)
      this.timers.push(timeoutId)
    }
    if (this.results != null) {
      resultIndex = this.results.length
      this.results[resultIndex] = null
    }
    this.pending++
    this.dispatchEvent(new QueueEvent('start', { job }))
    job.promise = job(next)
    if (job.promise !== undefined && typeof job.promise.then === 'function') {
      job.promise.then(function (result) {
        return next(undefined, result)
      }).catch(function (err) {
        return next(err || true)
      })
    }
    if (this.running && this.jobs.length > 0) {
      this._start()
    }
  }

  stop () {
    this.running = false
  }

  end (error) {
    this.clearTimers()
    this.jobs.length = 0
    this.pending = 0
    this.done(error)
  }

  clearTimers () {
    this.timers.forEach(timer => {
      clearTimeout(timer)
    })
    this.timers = []
  }

  _addCallbackToEndEvent (cb) {
    const onend = evt => {
      this.removeEventListener('end', onend)
      cb(evt.detail.error, this.results)
    }
    this.addEventListener('end', onend)
  }

  _createPromiseToEndEvent () {
    return new Promise((resolve, reject) => {
      this._addCallbackToEndEvent((error, results) => {
        if (error) reject(error)
        else resolve(results)
      })
    })
  }

  done (error) {
    this.session++
    this.running = false
    this.dispatchEvent(new QueueEvent('end', { error }))
  }
}

// Update event counts display in UI
export function updateEventCountsDisplay() {
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  const countsElement = document.getElementById('eventCounts')
  if (countsElement) {
    countsElement.textContent = `Commands: ${state.eventCounts.commands}, ` +
      `Launch Events: ${state.eventCounts.launchEvents}, ` +
      `Item Changes: ${state.eventCounts.itemChanges}`
    console.log('Global Data', state.globalData)
  }
}

// Initialize taskpane UI elements
export function initializeTaskpaneUI() {
  const statusElement = document.getElementById('status')
  if (statusElement) {
    const addinInstance = createAddIn()
    const hasItem = addinInstance.Office.context.mailbox && addinInstance.Office.context.mailbox.item
    if (hasItem) {
      statusElement.textContent = 'Aladdin is ready! Item selected.'
    } else {
      statusElement.textContent = 'Aladdin is ready! No item selected.'
    }
  }
  updateEventCountsDisplay()
}

// Show the taskpane programmatically
export function showAsTaskpane() {
  const addinInstance = createAddIn()

  if (!addinInstance.Office.addin) {
    console.error('Office.addin not available')
    return Promise.reject(new Error('Office.addin not available'))
  }

  if (typeof addinInstance.Office.addin.showAsTaskpane !== 'function') {
    console.error('Office.addin.showAsTaskpane not available')
    return Promise.reject(new Error('Office.addin.showAsTaskpane not available'))
  }

  return addinInstance.Office.addin.showAsTaskpane()
    .then(() => {
      console.log('Taskpane shown successfully')
      addinInstance.changeState({
        globalData: {
          lastTaskpaneShow: new Date().toISOString()
        }
      })
      return true
    })
    .catch(error => {
      console.error('Error showing taskpane:', error)
      throw error
    })
}

// Register ItemChanged event handler
export function registerItemChangedHandler() {
  const addinInstance = createAddIn()
  if (addinInstance.Office.context.mailbox && addinInstance.Office.context.mailbox.addHandlerAsync) {
    addinInstance.Office.context.mailbox.addHandlerAsync(
      addinInstance.Office.EventType.ItemChanged,
      onItemChanged,
      (asyncResult) => {
        if (asyncResult.status === addinInstance.Office.AsyncResultStatus.Failed) {
          console.error('Failed to register ItemChanged handler:', asyncResult.error.message)
        } else {
          console.log('ItemChanged handler registered successfully')
        }
      }
    )
  }
}

// ItemChanged event handler
export function onItemChanged(eventArgs) {
  console.log('ItemChanged event triggered', eventArgs)
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  addinInstance.changeState({
    eventCounts: {
      itemChanges: state.eventCounts.itemChanges + 1
    }
  })

  const hasItem = addinInstance.Office.context.mailbox && addinInstance.Office.context.mailbox.item
  // Update UI
  const statusElement = document.getElementById('status')
  if (statusElement) {
    if (hasItem) {
      const subject = addinInstance.Office.context.mailbox.item.subject || 'No subject'
      statusElement.textContent = `Item: ${subject}`
    } else {
      statusElement.textContent = 'Aladdin is ready! No item selected.'
    }
  }
  updateEventCountsDisplay()
}

// Register VisibilityChanged event handler
export function registerVisibilityChangedHandler() {
  const addinInstance = createAddIn()

  // Check if Office.addin exists
  if (!addinInstance.Office.addin) {
    console.warn('Office.addin not available')
    registerDocumentVisibilityHandler()
    return
  }

  // Try using setStartupBehavior if available (for visibility on startup)
  if (typeof addinInstance.Office.addin.setStartupBehavior === 'function') {
    addinInstance.Office.addin.setStartupBehavior(
      addinInstance.Office.StartupBehavior.load
    ).then(() => {
      console.log('Startup behavior set to load')
    }).catch(err => {
      console.warn('Could not set startup behavior:', err)
    })
  }

  // Try to register onVisibilityModeChanged
  if (typeof addinInstance.Office.addin.onVisibilityModeChanged === 'function') {
    try {
      addinInstance.Office.addin.onVisibilityModeChanged(onVisibilityChanged)
      console.log('Office.addin.onVisibilityModeChanged handler registered successfully')
    } catch (error) {
      console.error('Error registering onVisibilityModeChanged handler:', error)
      registerDocumentVisibilityHandler()
    }
  } else {
    console.warn('Office.addin.onVisibilityModeChanged not available, using fallback')
    registerDocumentVisibilityHandler()
  }
}

// Register document visibility change handler (fallback)
function registerDocumentVisibilityHandler() {
  if (typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', onDocumentVisibilityChanged)
    console.log('Document visibilitychange handler registered (fallback)')
  }
}

// Office.addin visibility change handler
export function onVisibilityChanged(args) {
  console.log('Office VisibilityChanged event triggered', args)
  const addinInstance = createAddIn()
  addinInstance.changeState({
    globalData: {
      lastVisibilityMode: args.visibilityMode,
      lastVisibilityChange: new Date().toISOString()
    }
  })

  const statusElement = document.getElementById('status')
  if (statusElement) {
    const mode = args.visibilityMode === addinInstance.Office.VisibilityMode.Hidden
      ? 'hidden'
      : 'visible'
    statusElement.textContent = `Taskpane is now ${mode}`
  }
  updateEventCountsDisplay()
}

// Document visibility change handler (fallback)
function onDocumentVisibilityChanged() {
  const addinInstance = createAddIn()
  const isHidden = document.hidden

  console.log('Document visibility changed, hidden:', isHidden)

  addinInstance.changeState({
    globalData: {
      lastVisibilityState: isHidden ? 'hidden' : 'visible',
      lastVisibilityChange: new Date().toISOString()
    }
  })

  const statusElement = document.getElementById('status')
  if (statusElement) {
    const mode = isHidden ? 'hidden' : 'visible'
    statusElement.textContent = `Taskpane is now ${mode}`
  }
  updateEventCountsDisplay()
}

// Command function for ribbon button
export function action(event) {
  console.log('Action command executed')
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  addinInstance.changeState({
    eventCounts: {
      commands: state.eventCounts.commands + 1
    },
    globalData: {
      lastAction: new Date().toISOString()
    }
  })

  updateEventCountsDisplay()
  event.completed()
}

// Handler for OnNewMessageCompose event
export function onNewMessageComposeHandler(event) {
  console.log('OnNewMessageCompose event triggered')
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  addinInstance.changeState({
    eventCounts: {
      launchEvents: state.eventCounts.launchEvents + 1
    },
    globalData: {
      lastEvent: 'OnNewMessageCompose'
    }
  })

  // Show taskpane when new message is composed
  showAsTaskpane()
    .then(() => {
      console.log('Taskpane opened from OnNewMessageCompose')
    })
    .catch(err => {
      console.warn('Could not show taskpane:', err)
    })

  updateEventCountsDisplay()
  event.completed()
}

// Handler for OnMessageSend event
export function onMessageSendHandler(event) {
  console.log('OnMessageSend event triggered')
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  addinInstance.changeState({
    eventCounts: {
      launchEvents: state.eventCounts.launchEvents + 1
    },
    globalData: {
      lastEvent: 'OnMessageSend'
    }
  })

  updateEventCountsDisplay()
  event.completed({ allowEvent: true })
}

// Handler for OnMessageRecipientsChanged event
export function onRecipientsChangedHandler(event) {
  console.log('OnMessageRecipientsChanged event triggered')
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  addinInstance.changeState({
    eventCounts: {
      launchEvents: state.eventCounts.launchEvents + 1
    },
    globalData: {
      lastEvent: 'OnMessageRecipientsChanged'
    }
  })

  updateEventCountsDisplay()
  event.completed()
}

// Handler for OnMessageFromChanged event
export function onFromChangedHandler(event) {
  console.log('OnMessageFromChanged event triggered')
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  addinInstance.changeState({
    eventCounts: {
      launchEvents: state.eventCounts.launchEvents + 1
    },
    globalData: {
      lastEvent: 'OnMessageFromChanged'
    }
  })

  updateEventCountsDisplay()
  event.completed()
}

// Initialize Office.actions associations
export function initializeAssociations(Office) {
  if (Office && Office.actions) {
    Office.actions.associate("action", action)
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler)
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler)
    Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler)
    Office.actions.associate("onFromChangedHandler", onFromChangedHandler)
    console.log('Office.actions associations registered')
  } else {
    console.warn('Office.actions not available for registration')
  }
}

// Initialize the add-in - called by Office.onReady
export function initializeAddIn(Office) {

  const addinInstance = createAddIn(Office)

  // Initialize taskpane UI if DOM is ready
  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initializeTaskpaneUI)
    } else {
      initializeTaskpaneUI()
    }
  }

  addinInstance.queue().push(cb => {
    const result = 'addin-initialized'
    cb(null, result)
  })
  addinInstance.start()

  registerItemChangedHandler()
  registerVisibilityChangedHandler()

  // Show taskpane after a delay
  setTimeout(() => {
    showAsTaskpane()
      .then(() => {
        console.log('Taskpane auto-opened after initialization')
      })
      .catch(err => {
        console.warn('Could not auto-open taskpane:', err)
      })
  }, 2000)

  return addinInstance
}