// taskpane.js - Taskpane execution context

import { initializeAddIn } from './addin.js'

let addinInstance = null;

Office.onReady((info) => {
  console.log('Office.onReady called in taskpane', info)

  if (info.host === Office.HostType.Outlook) {
    addinInstance = initializeAddIn('Taskpane')

    // Update UI to show ready status
    const statusElement = document.getElementById('status')
    if (statusElement) {
      statusElement.textContent = 'Aladdin is ready!'
    }

    addinInstance.queue().push(cb => {
      console.log('taskpane - Here')
      const result = 'taskpane'
      cb(null, result)
    })
  }
})