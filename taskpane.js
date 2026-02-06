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
    addinInstance.start()

    // Register ItemChanged event handler
    registerItemChangedHandler()
  }
})

function registerItemChangedHandler() {
  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      onItemChanged,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error('Failed to register ItemChanged handler:', asyncResult.error.message)
        } else {
          console.log('ItemChanged handler registered successfully')
        }
      }
    )
  } else {
    console.warn('addHandlerAsync not available in this context')
  }
}

function onItemChanged(eventArgs) {
  console.log('ItemChanged event triggered', eventArgs)

  if (addinInstance) {
    addinInstance.queue().push(cb => {
      console.log('Processing item change in queue')
      const result = 'item-changed'
      cb(null, result)
    })
    addinInstance.start()
  }

  // Update UI to reflect the item change
  const statusElement = document.getElementById('status')
  if (statusElement) {
    const itemType = Office.context.mailbox.item ? 'Item selected' : 'No item selected'
    statusElement.textContent = `${itemType} - Aladdin is ready!`
  }
}