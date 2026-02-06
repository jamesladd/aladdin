// runtime.js - Runtime initialization

import {
  initializeAddIn,
  initializeAssociations
} from './addin.js'

// Wait for Office to be available
function initializeRuntime() {
  if (typeof Office === 'undefined') {
    console.error('Office is not defined - Office.js may not be loaded')
    return
  }
  // Office.onReady initialization
  Office.onReady((info) => {
    console.log('Office.onReady called', info)
    if (info.host === Office.HostType.Outlook) {
      initializeAssociations(Office)
      initializeAddIn(Office)
    }
  })
}

// Call initialization
initializeRuntime()