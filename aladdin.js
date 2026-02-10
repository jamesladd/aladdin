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
    },
    state() {
      // return the current state
      return this._state
    },
    saveState() {
      // save the current state to local storage
    },
    loadState() {
      // load the current state from local storage
    },
    watchState() {
      // listen to local storage to know when it changed
    },
    event(name, details) {
      // record event named 'name' in state and store
    },
    initialize() {
    }
  }
}
