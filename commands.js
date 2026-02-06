// commands.js - Commands execution context (FunctionFile)

import { initializeAddIn } from './addin.js'

let addinInstance = null

Office.onReady((info) => {
  console.log('Office.onReady called in commands', info)

  if (info.host === Office.HostType.Outlook) {
    addinInstance = initializeAddIn('Commands')

    // Listen for ready event
    addinInstance.queue().addEventListener('ready', (e) => {
      console.log('Commands ready event received:', e.detail)
    })
  }
})

// Command function that gets called when the ribbon button is clicked
function action(event) {
  console.log('Action command executed')

  if (addinInstance) {
    console.log('Add-in instance is available in command context')
  }

  // Signal that the command is complete
  event.completed()
}

// Register the function
Office.actions.associate("action", action)