// addin.js

export function createAddIn(Office) {
  if (typeof window !== 'undefined' && window.aladdinInstance) {
    console.log('createAddIn - instance from window')
    return window.aladdinInstance
  }

  console.log('createAddIn - new')
  const queue = new Queue()
  const instance = addin(queue, Office)
  if (typeof window !== 'undefined') window.aladdinInstance = instance;
  return instance
}

function addin(queue, Office) {
  return {
    queue() {
      return queue
    },
    start() {
      queue.addEventListener('start', e => {
        console.log('Jobs start:', e)
      })
      queue.addEventListener('success', e => {
        console.log('Job ok:', e)
      })
      queue.addEventListener('error', e => {
        console.log('Job err:', e)
      })
      queue.addEventListener('end', e => {
        console.log('Jobs end:', e)
      })
      queue.start(err => {
        if (err) console.error(err)
      })
    },
    Office,
    globalData: {},
    eventCounts: {
      commands: 0,
      launchEvents: 0,
      itemChanges: 0
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
  const countsElement = document.getElementById('eventCounts')
  if (countsElement) {
    countsElement.textContent = `Commands: ${addinInstance.eventCounts.commands}, ` +
      `Launch Events: ${addinInstance.eventCounts.launchEvents}, ` +
      `Item Changes: ${addinInstance.eventCounts.itemChanges}`
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
  addinInstance.eventCounts.itemChanges++
  const hasItem = addinInstance.Office.context.mailbox && addinInstance.Office.context.mailbox.item

  addinInstance.queue().push(cb => {
    console.log('Processing item change in shared queue')
    const result = hasItem ? 'item-changed-with-item' : 'item-changed-no-item'
    cb(null, result)
  })

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

  addinInstance.start()
  updateEventCountsDisplay()
}

// Command function for ribbon button
export function action(event) {
  console.log('Action command executed')

  const addinInstance = createAddIn()
  addinInstance.eventCounts.commands++
  addinInstance.globalData.lastAction = new Date().toISOString()

  console.log('Add-in instance is available')
  console.log('Global data:', addinInstance.globalData)

  addinInstance.queue().push(cb => {
    console.log('Action queued')
    const result = 'action-command'
    cb(null, result)
  })

  addinInstance.start()
  updateEventCountsDisplay()

  event.completed()
}

// Handler for OnNewMessageCompose event
export function onNewMessageComposeHandler(event) {
  console.log('OnNewMessageCompose event triggered')

  const addinInstance = createAddIn()
  addinInstance.eventCounts.launchEvents++
  addinInstance.globalData.lastEvent = 'OnNewMessageCompose'

  console.log('Shared state accessible in launch event')
  console.log('Event counts:', addinInstance.eventCounts)

  addinInstance.queue().push(cb => {
    console.log('OnNewMessageCompose queued')
    cb(null, 'new-message-compose')
  })

  addinInstance.start()
  updateEventCountsDisplay()

  event.completed()
}

// Handler for OnMessageSend event
export function onMessageSendHandler(event) {
  console.log('OnMessageSend event triggered')

  const addinInstance = createAddIn()
  addinInstance.eventCounts.launchEvents++
  addinInstance.globalData.lastEvent = 'OnMessageSend'

  addinInstance.queue().push(cb => {
    console.log('OnMessageSend queued')
    cb(null, 'message-send')
  })

  addinInstance.start()
  updateEventCountsDisplay()

  event.completed({ allowEvent: true })
}

// Handler for OnMessageRecipientsChanged event
export function onRecipientsChangedHandler(event) {
  console.log('OnMessageRecipientsChanged event triggered')

  const addinInstance = createAddIn()
  addinInstance.eventCounts.launchEvents++
  addinInstance.globalData.lastEvent = 'OnMessageRecipientsChanged'

  addinInstance.queue().push(cb => {
    console.log('OnRecipientsChanged queued')
    cb(null, 'recipients-changed')
  })

  addinInstance.start()
  updateEventCountsDisplay()

  event.completed()
}

// Handler for OnMessageFromChanged event
export function onFromChangedHandler(event) {
  console.log('OnMessageFromChanged event triggered')

  const addinInstance = createAddIn()
  addinInstance.eventCounts.launchEvents++
  addinInstance.globalData.lastEvent = 'OnMessageFromChanged'

  addinInstance.queue().push(cb => {
    console.log('OnFromChanged queued')
    cb(null, 'from-changed')
  })

  addinInstance.start()
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
    console.log('Add-in initialized')
    const result = 'addin-initialized'
    cb(null, result)
  })
  addinInstance.start()

  registerItemChangedHandler()

  return addinInstance
}