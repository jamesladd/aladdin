// aladdin.js

const singleton = [false]
const STATE_KEY = 'aladdin-addin-state'
const MAX_ITEM_HISTORY = 10

export function createAddIn(Office) {
  if (typeof window !== 'undefined' && window.aladdinInstance) return window.aladdinInstance;
  if (singleton[0]) return singleton[0];
  const queue = new Queue()
  const instance = aladdin(queue, Office)
  if (typeof window !== 'undefined') window.aladdinInstance = instance;
  if (typeof window === 'undefined') singleton[0] = instance;
  instance.loadState()
  instance.watchStorage()
  return instance
}

function aladdin(queue, Office) {
  return {
    queue() {
      return queue
    },
    start() {
      // queue.addEventListener('success', e => {
      //   console.log('Job ok:', JSON.stringify(e.detail, null, 2))
      // })
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
        itemChanges: 0,
        changeChecks: 0
      },
      itemHistory: {}, // Store historical snapshots by itemId (limited to MAX_ITEM_HISTORY)
      currentItem: null // Store the currently selected item's snapshot
    },
    _storageWatcher: null,
    _pollInterval: null,
    state() {
      return this._state
    },
    saveState() {
      if (typeof localStorage !== 'undefined') {
        try {
          const stateJson = JSON.stringify(this._state)
          localStorage.setItem(STATE_KEY, stateJson)
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
            const loadedState = JSON.parse(stateJson)
            // Ensure itemHistory and currentItem exist even in old saved states
            if (!loadedState.itemHistory) {
              loadedState.itemHistory = {}
            }
            if (!loadedState.currentItem) {
              loadedState.currentItem = null
            }
            if (!loadedState.eventCounts.changeChecks) {
              loadedState.eventCounts.changeChecks = 0
            }
            this._state = loadedState
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
      if (changes.itemHistory) Object.assign(this._state.itemHistory, changes.itemHistory);
      if (changes.hasOwnProperty('currentItem')) this._state.currentItem = changes.currentItem;
      this.saveState()
    },
    watchStorage() {
      if (typeof window !== 'undefined' && typeof window.addEventListener === 'function') {
        // Listen for storage events (from other tabs/windows)
        this._storageWatcher = (event) => {
          if (event.key === STATE_KEY && event.newValue) {
            try {
              this._state = JSON.parse(event.newValue)
              this.onStateChanged()
            } catch (error) {
              console.error('Error parsing external state change:', error)
            }
          }
        }
        window.addEventListener('storage', this._storageWatcher)

        // Also poll for changes in same window (storage event doesn't fire for same-window changes)
        this._pollInterval = setInterval(() => {
          this.checkStateChanged()
        }, 2000)
      }
    },
    unwatchStorage() {
      if (this._storageWatcher && typeof window !== 'undefined') {
        window.removeEventListener('storage', this._storageWatcher)
        this._storageWatcher = null
      }
      if (this._pollInterval) {
        clearInterval(this._pollInterval)
        this._pollInterval = null
      }
    },
    checkStateChanged() {
      if (typeof localStorage !== 'undefined') {
        try {
          const stateJson = localStorage.getItem(STATE_KEY)
          if (stateJson) {
            const newState = JSON.parse(stateJson)
            // Simple comparison - in production you might want deep equality check
            if (JSON.stringify(newState) !== JSON.stringify(this._state)) {
              this._state = newState
              this.onStateChanged()
            }
          }
        } catch (error) {
          console.error('Error checking state changes:', error)
        }
      }
    },
    onStateChanged() {
      console.log('State changed:', this._state)
      // Update UI when state changes
      updateEventCountsDisplay()
    },
    cleanup() {
      this.unwatchStorage()
      this.queue().stop()
      cleanupEventListeners()
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

// Platform detection
export function isDesktop() {
  if (typeof Office === 'undefined') return false

  // Check if running in desktop Outlook
  const isDesktopPlatform = Office.context?.platform === Office.PlatformType.PC ||
    Office.context?.platform === Office.PlatformType.Mac

  // Desktop typically doesn't have Office.addin or has limited support
  const hasFullAddinAPI = typeof Office.addin !== 'undefined' &&
    typeof Office.addin.showAsTaskpane === 'function'

  return isDesktopPlatform || !hasFullAddinAPI
}

// Limit itemHistory to MAX_ITEM_HISTORY items (keep most recent)
function limitItemHistory(itemHistory) {
  const entries = Object.entries(itemHistory);

  if (entries.length <= MAX_ITEM_HISTORY) {
    return itemHistory;
  }

  // Sort by timestamp (most recent first)
  entries.sort((a, b) => {
    const timeA = new Date(a[1].timestamp).getTime();
    const timeB = new Date(b[1].timestamp).getTime();
    return timeB - timeA;
  });

  // Keep only the most recent MAX_ITEM_HISTORY items
  const limitedEntries = entries.slice(0, MAX_ITEM_HISTORY);

  return Object.fromEntries(limitedEntries);
}

// Capture a snapshot of current email item
export function captureItemSnapshot(item) {
  if (!item) return Promise.resolve(null);

  return new Promise((resolve) => {
    const snapshot = {
      itemId: item.itemId,
      conversationId: item.conversationId,
      subject: item.subject || '',
      timestamp: new Date().toISOString()
    };

    // Capture categories
    item.categories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        snapshot.categories = result.value || [];
      } else {
        snapshot.categories = [];
      }

      // Capture item class (gives context about item type and location)
      snapshot.itemClass = item.itemClass || '';

      // Try to get folder information if available (limited API support)
      if (item.getItemIdAsync) {
        item.getItemIdAsync((folderResult) => {
          if (folderResult.status === Office.AsyncResultStatus.Succeeded) {
            snapshot.folderId = folderResult.value;
          }
          continueCapture();
        });
      } else {
        continueCapture();
      }

      function continueCapture() {
        // Capture from address
        if (item.from) {
          snapshot.from = {
            displayName: item.from.displayName || '',
            emailAddress: item.from.emailAddress || ''
          };
        }

        // Capture to recipients
        if (item.to) {
          item.to.getAsync((toResult) => {
            if (toResult.status === Office.AsyncResultStatus.Succeeded) {
              snapshot.to = toResult.value.map(r => ({
                displayName: r.displayName || '',
                emailAddress: r.emailAddress || ''
              }));
            } else {
              snapshot.to = [];
            }

            // Capture cc recipients
            if (item.cc) {
              item.cc.getAsync((ccResult) => {
                if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                  snapshot.cc = ccResult.value.map(r => ({
                    displayName: r.displayName || '',
                    emailAddress: r.emailAddress || ''
                  }));
                } else {
                  snapshot.cc = [];
                }
                resolve(snapshot);
              });
            } else {
              snapshot.cc = [];
              resolve(snapshot);
            }
          });
        } else {
          snapshot.to = [];
          snapshot.cc = [];
          resolve(snapshot);
        }
      }
    });
  });
}

// Re-read an item using REST API to get current state
export function rereadItemSnapshot(itemId, Office) {
  return new Promise((resolve, reject) => {
    // Office.js doesn't provide a direct way to read an item by ID when it's not selected
    // We need to use REST API or EWS

    // Get access token for REST API
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error('Failed to get access token for REST API:', result.error);
        reject(new Error('Failed to get access token'));
        return;
      }

      const accessToken = result.value;
      const restUrl = Office.context.mailbox.restUrl;

      // Convert itemId to REST format
      const restId = Office.context.mailbox.convertToRestId(
        itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );

      // Construct the REST API URL
      const getMessageUrl = `${restUrl}/v2.0/me/messages/${restId}?$select=subject,categories,from,toRecipients,ccRecipients,parentFolderId,itemClass`;

      // Make REST API call
      fetch(getMessageUrl, {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        }
      })
        .then(response => {
          if (!response.ok) {
            throw new Error(`REST API returned ${response.status}: ${response.statusText}`);
          }
          return response.json();
        })
        .then(data => {
          // Convert REST API response to our snapshot format
          const snapshot = {
            itemId: itemId,
            conversationId: data.conversationId || '',
            subject: data.subject || '',
            categories: data.categories || [],
            itemClass: data.itemClass || '',
            folderId: data.parentFolderId || '',
            timestamp: new Date().toISOString()
          };

          // Convert from address
          if (data.from && data.from.emailAddress) {
            snapshot.from = {
              displayName: data.from.emailAddress.name || '',
              emailAddress: data.from.emailAddress.address || ''
            };
          }

          // Convert to recipients
          if (data.toRecipients) {
            snapshot.to = data.toRecipients.map(r => ({
              displayName: r.emailAddress.name || '',
              emailAddress: r.emailAddress.address || ''
            }));
          } else {
            snapshot.to = [];
          }

          // Convert cc recipients
          if (data.ccRecipients) {
            snapshot.cc = data.ccRecipients.map(r => ({
              displayName: r.emailAddress.name || '',
              emailAddress: r.emailAddress.address || ''
            }));
          } else {
            snapshot.cc = [];
          }

          resolve(snapshot);
        })
        .catch(error => {
          console.error('Error fetching item via REST API:', error);
          reject(error);
        });
    });
  });
}

// Compare two item snapshots and return detected changes
export function compareItemSnapshots(oldSnapshot, newSnapshot) {
  if (!oldSnapshot || !newSnapshot) return null;

  const changes = {
    itemId: newSnapshot.itemId,
    timestamp: new Date().toISOString(),
    detected: []
  };

  // Check category changes
  const oldCategories = (oldSnapshot.categories || []).sort().join(',');
  const newCategories = (newSnapshot.categories || []).sort().join(',');
  if (oldCategories !== newCategories) {
    changes.detected.push({
      type: 'categories',
      old: oldSnapshot.categories || [],
      new: newSnapshot.categories || []
    });
  }

  // Check folder/location changes (using folderId)
  if (oldSnapshot.folderId !== newSnapshot.folderId && oldSnapshot.folderId && newSnapshot.folderId) {
    changes.detected.push({
      type: 'folder',
      old: oldSnapshot.folderId,
      new: newSnapshot.folderId
    });
  }

  // Also check itemClass changes (can indicate moves or type changes)
  if (oldSnapshot.itemClass !== newSnapshot.itemClass) {
    changes.detected.push({
      type: 'itemClass',
      old: oldSnapshot.itemClass,
      new: newSnapshot.itemClass
    });
  }

  // Check from address changes
  const oldFrom = oldSnapshot.from ? oldSnapshot.from.emailAddress : '';
  const newFrom = newSnapshot.from ? newSnapshot.from.emailAddress : '';
  if (oldFrom !== newFrom) {
    changes.detected.push({
      type: 'from',
      old: oldSnapshot.from,
      new: newSnapshot.from
    });
  }

  // Check to recipients changes
  const oldTo = (oldSnapshot.to || []).map(r => r.emailAddress).sort().join(',');
  const newTo = (newSnapshot.to || []).map(r => r.emailAddress).sort().join(',');
  if (oldTo !== newTo) {
    changes.detected.push({
      type: 'to',
      old: oldSnapshot.to || [],
      new: newSnapshot.to || []
    });
  }

  // Check cc recipients changes
  const oldCc = (oldSnapshot.cc || []).map(r => r.emailAddress).sort().join(',');
  const newCc = (newSnapshot.cc || []).map(r => r.emailAddress).sort().join(',');
  if (oldCc !== newCc) {
    changes.detected.push({
      type: 'cc',
      old: oldSnapshot.cc || [],
      new: newSnapshot.cc || []
    });
  }

  return changes.detected.length > 0 ? changes : null;
}

// Display detected item changes in UI
export function displayItemChanges(changes) {
  const changesElement = document.getElementById('itemChanges');
  if (!changesElement) return;

  let html = '<div class="changes-header">Changes Detected:</div>';

  changes.detected.forEach(change => {
    switch(change.type) {
      case 'categories':
        html += `<div class="change-item">
          <strong>Categories:</strong><br>
          Old: ${change.old.join(', ') || '(none)'}<br>
          New: ${change.new.join(', ') || '(none)'}
        </div>`;
        break;

      case 'folder':
        html += `<div class="change-item">
          <strong>Folder Changed</strong><br>
          Item was moved to a different folder
        </div>`;
        break;

      case 'itemClass':
        html += `<div class="change-item">
          <strong>Item Class Changed</strong><br>
          Old: ${change.old || '(unknown)'}<br>
          New: ${change.new || '(unknown)'}
        </div>`;
        break;

      case 'from':
        html += `<div class="change-item">
          <strong>From:</strong><br>
          Old: ${change.old ? change.old.emailAddress : '(none)'}<br>
          New: ${change.new ? change.new.emailAddress : '(none)'}
        </div>`;
        break;

      case 'to':
        html += `<div class="change-item">
          <strong>To Recipients:</strong><br>
          Old: ${change.old.map(r => r.emailAddress).join(', ') || '(none)'}<br>
          New: ${change.new.map(r => r.emailAddress).join(', ') || '(none)'}
        </div>`;
        break;

      case 'cc':
        html += `<div class="change-item">
          <strong>CC Recipients:</strong><br>
          Old: ${change.old.map(r => r.emailAddress).join(', ') || '(none)'}<br>
          New: ${change.new.map(r => r.emailAddress).join(', ') || '(none)'}
        </div>`;
        break;
    }
  });

  changesElement.innerHTML = html;
}

// Debounce timer for change checks
let checkDebounceTimer = null;

// Check current item for changes (without switching away)
export function checkCurrentItemForChanges() {
  const addinInstance = createAddIn();
  const Office = addinInstance.Office;
  const state = addinInstance.state();
  const currentItem = state.currentItem;

  if (!currentItem) {
    console.log('No current item to check');
    return Promise.resolve(null);
  }

  const liveItem = Office.context.mailbox.item;
  if (!liveItem || liveItem.itemId !== currentItem.itemId) {
    console.log('Live item does not match current item');
    return Promise.resolve(null);
  }

  console.log('Checking current item for changes:', currentItem.itemId);

  // Increment changeChecks counter
  addinInstance.changeState({
    eventCounts: {
      changeChecks: state.eventCounts.changeChecks + 1
    }
  });

  // Capture fresh snapshot of currently selected item
  return captureItemSnapshot(liveItem).then(newSnapshot => {
    const changes = compareItemSnapshots(currentItem, newSnapshot);

    if (changes) {
      console.log('Changes detected in current item:', changes);

      // Update currentItem with new snapshot
      addinInstance.changeState({
        currentItem: newSnapshot,
        globalData: {
          lastCurrentItemChanges: changes,
          lastChangeCheckTime: new Date().toISOString()
        }
      });

      displayItemChanges(changes);
      updateEventCountsDisplay();

      return changes;
    } else {
      console.log('No changes detected in current item');

      // Update the timestamp even if no changes
      addinInstance.changeState({
        currentItem: newSnapshot,
        globalData: {
          lastChangeCheckTime: new Date().toISOString()
        }
      });

      // Show "no changes" message
      const changesElement = document.getElementById('itemChanges');
      if (changesElement) {
        changesElement.innerHTML = '<div class="info-message">No changes detected</div>';
      }

      updateEventCountsDisplay();

      return null;
    }
  }).catch(error => {
    console.error('Error checking current item for changes:', error);
    return null;
  });
}

// Debounced version to avoid excessive checks
export function debouncedCheckCurrentItem() {
  if (checkDebounceTimer) clearTimeout(checkDebounceTimer);
  checkDebounceTimer = setTimeout(() => {
    checkCurrentItemForChanges();
  }, 1000);
}

// Update event counts display in UI
export function updateEventCountsDisplay() {
  const addinInstance = createAddIn()
  const state = addinInstance.state()
  const countsElement = document.getElementById('eventCounts')
  if (countsElement) {
    countsElement.textContent = `Commands: ${state.eventCounts.commands}, ` +
      `Launch Events: ${state.eventCounts.launchEvents}, ` +
      `Item Changes: ${state.eventCounts.itemChanges}, ` +
      `Change Checks: ${state.eventCounts.changeChecks}`
  }
}

// Initialize taskpane UI elements
export function initializeTaskpaneUI() {
  const statusElement = document.getElementById('status')
  if (statusElement) {
    const addinInstance = createAddIn()
    const hasItem = addinInstance.Office.context.mailbox && addinInstance.Office.context.mailbox.item
    const platform = isDesktop() ? 'Desktop' : 'Web'
    if (hasItem) {
      statusElement.textContent = `Aladdin is ready on ${platform}! Item selected.`
    } else {
      statusElement.textContent = `Aladdin is ready on ${platform}! No item selected.`
    }
  }
  updateEventCountsDisplay()

  // Setup refresh button
  const refreshButton = document.getElementById('refreshButton')
  if (refreshButton) {
    refreshButton.addEventListener('click', () => {
      console.log('Refresh button clicked');
      checkCurrentItemForChanges();
    });
  }

  // Register event listeners for multi-event strategy
  registerMultiEventListeners();
}

// Store event listeners for cleanup
let eventListeners = [];

// Register multiple event listeners to trigger change checks
function registerMultiEventListeners() {
  console.log('Registering multi-event listeners');

  // Window focus event
  const focusHandler = () => {
    console.log('Window gained focus, checking for changes');
    debouncedCheckCurrentItem();
  };
  window.addEventListener('focus', focusHandler);
  eventListeners.push({ target: window, event: 'focus', handler: focusHandler });

  // Document visibility change
  const visibilityHandler = () => {
    if (!document.hidden) {
      console.log('Document became visible, checking for changes');
      debouncedCheckCurrentItem();
    }
  };
  document.addEventListener('visibilitychange', visibilityHandler);
  eventListeners.push({ target: document, event: 'visibilitychange', handler: visibilityHandler });

  // User clicks in taskpane
  const clickHandler = () => {
    debouncedCheckCurrentItem();
  };
  document.addEventListener('click', clickHandler);
  eventListeners.push({ target: document, event: 'click', handler: clickHandler });

  // User presses key in taskpane
  const keydownHandler = () => {
    debouncedCheckCurrentItem();
  };
  document.addEventListener('keydown', keydownHandler);
  eventListeners.push({ target: document, event: 'keydown', handler: keydownHandler });

  console.log('Multi-event listeners registered');
}

// Cleanup event listeners
function cleanupEventListeners() {
  console.log('Cleaning up event listeners');
  eventListeners.forEach(({ target, event, handler }) => {
    target.removeEventListener(event, handler);
  });
  eventListeners = [];
}

// Show the taskpane programmatically
export function showAsTaskpane() {
  const addinInstance = createAddIn()

  // Check if we're on desktop - if so, don't try to show programmatically
  if (isDesktop()) {
    console.log('Desktop Outlook detected - taskpane must be opened manually via ribbon')
    return Promise.resolve(false)
  }

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

// ItemChanged event handler - CORRECTED LOGIC
export function onItemChanged(eventArgs) {
  console.log('ItemChanged event triggered', eventArgs);
  const addinInstance = createAddIn();
  const Office = addinInstance.Office;
  const state = addinInstance.state();

  addinInstance.changeState({
    eventCounts: {
      itemChanges: state.eventCounts.itemChanges + 1
    }
  });

  // STEP 1: Check if we have a previous item to compare
  const previousItem = state.currentItem;

  if (previousItem) {
    console.log('Previous item detected, re-reading to check for changes:', previousItem.itemId);

    // STEP 2: Re-read the previous item to get its current state
    rereadItemSnapshot(previousItem.itemId, Office)
      .then(rereadSnapshot => {
        console.log('Re-read snapshot obtained:', rereadSnapshot);

        // STEP 3: Compare the original snapshot with the re-read one
        const changes = compareItemSnapshots(previousItem, rereadSnapshot);

        if (changes) {
          console.log('Changes detected in previous item:', changes);

          // Store detected changes
          addinInstance.changeState({
            globalData: {
              lastItemChanges: changes
            }
          });

          // Store the final state in history (with limit)
          const updatedHistory = { ...state.itemHistory };
          updatedHistory[previousItem.itemId] = rereadSnapshot;
          const limitedHistory = limitItemHistory(updatedHistory);

          addinInstance.changeState({
            itemHistory: limitedHistory
          });

          // Display changes in UI
          displayItemChanges(changes);
        } else {
          console.log('No changes detected in previous item');

          // Store unchanged snapshot in history (with limit)
          const updatedHistory = { ...state.itemHistory };
          updatedHistory[previousItem.itemId] = rereadSnapshot;
          const limitedHistory = limitItemHistory(updatedHistory);

          addinInstance.changeState({
            itemHistory: limitedHistory
          });

          // Clear changes display
          const changesElement = document.getElementById('itemChanges');
          if (changesElement) {
            changesElement.innerHTML = '<div class="info-message">No changes detected in previous item</div>';
          }
        }
      })
      .catch(error => {
        console.error('Error re-reading previous item:', error);

        // Even if re-read fails, store the original snapshot (with limit)
        const updatedHistory = { ...state.itemHistory };
        updatedHistory[previousItem.itemId] = previousItem;
        const limitedHistory = limitItemHistory(updatedHistory);

        addinInstance.changeState({
          itemHistory: limitedHistory
        });
      })
      .finally(() => {
        // STEP 4: Now capture the new current item
        captureNewCurrentItem();
      });
  } else {
    console.log('No previous item to check');
    // No previous item, just capture the new one
    captureNewCurrentItem();
  }

  // Helper function to capture the new current item
  function captureNewCurrentItem() {
    const newItem = Office.context.mailbox.item;

    if (newItem) {
      // Capture the new item's snapshot
      captureItemSnapshot(newItem).then(newSnapshot => {
        if (!newSnapshot) {
          console.error('Failed to capture new item snapshot');
          return;
        }

        console.log('New item snapshot captured:', newSnapshot);

        // Store as current item
        addinInstance.changeState({
          currentItem: newSnapshot
        });

        // Update UI with current item
        const subject = newItem.subject || 'No subject';
        const statusElement = document.getElementById('status');
        if (statusElement) {
          statusElement.textContent = `Item: ${subject}`;
        }
      }).catch(error => {
        console.error('Error capturing new item snapshot:', error);
      });
    } else {
      console.log('No new item selected');

      // Clear current item
      addinInstance.changeState({
        currentItem: null
      });

      const platform = isDesktop() ? 'Desktop' : 'Web';
      const statusElement = document.getElementById('status');
      if (statusElement) {
        statusElement.textContent = `Aladdin is ready on ${platform}! No item selected.`;
      }

      // Clear changes display
      const changesElement = document.getElementById('itemChanges');
      if (changesElement) {
        changesElement.innerHTML = '';
      }
    }
  }

  updateEventCountsDisplay();
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

  // Check for changes when taskpane becomes visible
  if (args.visibilityMode !== addinInstance.Office.VisibilityMode.Hidden) {
    console.log('Taskpane became visible, checking for changes');
    debouncedCheckCurrentItem();
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

  // Must call event.completed() for both desktop and web
  if (event && event.completed) {
    event.completed()
  }
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

  // Show taskpane when new message is composed (web only)
  if (!isDesktop()) {
    showAsTaskpane()
      .then(() => {
        console.log('Taskpane opened from OnNewMessageCompose')
      })
      .catch(err => {
        console.warn('Could not show taskpane:', err)
      })
  }

  updateEventCountsDisplay()

  if (event && event.completed) {
    event.completed()
  }
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

  if (event && event.completed) {
    event.completed({ allowEvent: true })
  }
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

  // Check for changes since we're in compose mode and recipients changed
  debouncedCheckCurrentItem()

  updateEventCountsDisplay()

  if (event && event.completed) {
    event.completed()
  }
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

  // Check for changes since we're in compose mode and from changed
  debouncedCheckCurrentItem()

  updateEventCountsDisplay()

  if (event && event.completed) {
    event.completed()
  }
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

    // Cleanup when window unloads
    window.addEventListener('beforeunload', () => {
      addinInstance.cleanup()
    })
    window.addEventListener('pagehide', () => {
      addinInstance.cleanup()
    })
  }

  addinInstance.queue().push(cb => {
    const result = 'addin-initialized'
    cb(null, result)
  })
  addinInstance.start()

  registerItemChangedHandler()
  registerVisibilityChangedHandler()

  // Capture initial item if one is selected
  const initialItem = Office.context.mailbox.item;
  if (initialItem) {
    console.log('Initial item detected, capturing snapshot');
    captureItemSnapshot(initialItem).then(snapshot => {
      if (snapshot) {
        addinInstance.changeState({
          currentItem: snapshot
        });
        console.log('Initial item snapshot captured:', snapshot);
      }
    }).catch(error => {
      console.error('Error capturing initial item:', error);
    });
  }

  // Only auto-show taskpane on web, not desktop
  if (!isDesktop()) {
    setTimeout(() => {
      showAsTaskpane()
        .then(() => {
          console.log('Taskpane auto-opened after initialization')
        })
        .catch(err => {
          console.warn('Could not auto-open taskpane:', err)
        })
    }, 2000)
  } else {
    console.log('Desktop Outlook detected - taskpane will not auto-open')
  }

  return addinInstance
}