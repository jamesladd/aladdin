// taskpane.js - Taskpane execution context

import { initializeAddIn } from './addin.js'

let addinInstance = null;

Office.onReady((info) => {
  console.log('Office.onReady called in taskpane', info)

  if (info.host === Office.HostType.Outlook) {
    addinInstance = initializeAddIn('Taskpane')

    // Check if an item is available
    const hasItem = Office.context.mailbox && Office.context.mailbox.item

    // Update UI to show ready status
    const statusElement = document.getElementById('status')
    if (statusElement) {
      if (hasItem) {
        statusElement.textContent = 'Aladdin is ready! Item selected.'
      } else {
        statusElement.textContent = 'Aladdin is ready! No item selected.'
      }
    }

    addinInstance.queue().push(cb => {
      console.log('taskpane - Here')
      const result = 'taskpane'
      cb(null, result)
    })
    addinInstance.start()

    setInterval(_ => console.log('Interval'), 3000)

    // Register ItemChanged event handler only if supported
    if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
      registerItemChangedHandler()
    } else {
      console.log('ItemChanged handler not supported in this context')
    }
  }
})

function registerItemChangedHandler() {
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
}

function onItemChanged(eventArgs) {
  console.log('ItemChanged event triggered', eventArgs)

  const hasItem = Office.context.mailbox && Office.context.mailbox.item

  if (addinInstance) {
    addinInstance.queue().push(cb => {
      console.log('Processing item change in queue')
      const result = hasItem ? 'item-changed-with-item' : 'item-changed-no-item'
      cb(null, result)
    })
    addinInstance.start()
  }

  // Update UI to reflect the item change
  const statusElement = document.getElementById('status')
  if (statusElement) {
    if (hasItem) {
      const subject = Office.context.mailbox.item.subject || 'No subject'
      statusElement.textContent = `Item: ${subject}`
    } else {
      statusElement.textContent = 'Aladdin is ready! No item selected.'
    }
  }
}