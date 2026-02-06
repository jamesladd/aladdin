// runtime.js - Runtime initialization

import {
  initializeAddIn,
  initializeAssociations
} from './addin.js'

// Office.onReady initialization
Office.onReady((info) => {
  console.log('Office.onReady called', info)

  if (info.host === Office.HostType.Outlook) {
    // Initialize associations first
    initializeAssociations(Office)

    // Then initialize the add-in
    initializeAddIn(Office)
  }
})