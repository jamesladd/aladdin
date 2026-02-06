// launchevent.js - Direct JavaScript runtime for LaunchEvents

import { initializeAddIn } from './addin.js'

let addinInstance = null;

(function() {
  console.log('Aladin - launchevent')
})();

// Initialize when the runtime loads
Office.onReady((info) => {
  console.log('Office.onReady called in launchevent', info)

  if (info.host === Office.HostType.Outlook) {
    addinInstance = initializeAddIn('LaunchEvent')
    addinInstance.queue().push(cb => {
      const result = 'launchevent'
      cb(null, result)
    })
  }
})

// Handler for OnNewMessageCompose event
function onNewMessageComposeHandler(event) {
  console.log('OnNewMessageCompose event triggered')

  if (addinInstance) {
    console.log('Add-in instance is available in launch event context')
  }

  // Signal that the event handler is complete
  event.completed()
}

// Handler for OnMessageSend event
function onMessageSendHandler(event) {
  console.log('OnMessageSend event triggered')

  if (addinInstance) {
    console.log('Add-in instance is available in launch event context')
  }

  // Allow the message to be sent
  event.completed({ allowEvent: true })
}

// Register the functions with Office
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler)
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler)
}