// runtime.js - Runtime initialization

import { createAladdin } from './aladdin.js'

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
      const addin = createAladdin(Office)

      // Gather initialization info
      const userName = Office.context.mailbox.userProfile.displayName || 'Unknown User'

      // Folder: not directly available, use default; will be updated from headers if possible
      const folderName = 'Inbox (default)'

      const platform = Office.context.platform || 'Unknown'

      let version = 'Unknown'
      if (Office.context.diagnostics && Office.context.diagnostics.version) {
        version = Office.context.diagnostics.version
      } else if (Office.context.mailbox.diagnostics && Office.context.mailbox.diagnostics.version) {
        version = Office.context.mailbox.diagnostics.version
      }

      addin.initialize(userName, folderName, platform, version)
    }
  })
}

// Call initialization
initializeRuntime()