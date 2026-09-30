/**
 * DSH Bridge Plugin for R7-Office Desktop Editors.
 * Connects active R7 Desktop window to DSH MCP Server via local WebSocket.
 */
(function(window, undefined) {
  let ws = null
  let reconnectTimer = null
  const WS_URL = 'ws://127.0.0.1:7888/r7-bridge'

  window.Asc.plugin.init = function() {
    connectBridge()
  }

  function connectBridge() {
    try {
      ws = new WebSocket(WS_URL)

      ws.onopen = function() {
        if (reconnectTimer) {
          clearTimeout(reconnectTimer)
          reconnectTimer = null
        }
        ws.send(JSON.stringify({
          type: 'register',
          editorType: window.Asc.plugin.info ? window.Asc.plugin.info.editorType : 'unknown',
          guid: window.Asc.plugin.info ? window.Asc.plugin.info.guid : 'asc.{D5B29457-194D-4E9A-A37F-02D739818FE1}'
        }))
      }

      ws.onmessage = function(event) {
        try {
          const msg = JSON.parse(event.data)
          handleCommand(msg)
        } catch (e) {
          console.error('[DSH Bridge] Message handling error:', e)
        }
      }

      ws.onclose = function() {
        scheduleReconnect()
      }

      ws.onerror = function() {
        scheduleReconnect()
      }
    } catch (e) {
      scheduleReconnect()
    }
  }

  function scheduleReconnect() {
    if (!reconnectTimer) {
      reconnectTimer = setTimeout(connectBridge, 3000)
    }
  }

  function handleCommand(msg) {
    const { id, action, payload } = msg

    if (action === 'ping') {
      sendReply(id, { pong: true, time: Date.now() })
      return
    }

    if (action === 'getSelection') {
      window.Asc.plugin.executeMethod('GetSelectedText', [], function(selectedText) {
        sendReply(id, { success: true, text: selectedText })
      })
      return
    }

    if (action === 'replaceSelection') {
      const textToPaste = payload ? payload.text : ''
      window.Asc.plugin.executeMethod('PasteText', [textToPaste], function() {
        sendReply(id, { success: true })
      })
      return
    }

    if (action === 'safeCommand') {
      const { command, args = {} } = payload || {}

      if (command === 'getSelectedText') {
        window.Asc.plugin.executeMethod('GetSelectedText', [], function(text) {
          sendReply(id, { success: true, text: text })
        })
        return
      }

      if (command === 'setSelectedText') {
        window.Asc.plugin.executeMethod('PasteText', [args.text || ''], function() {
          sendReply(id, { success: true })
        })
        return
      }

      if (command === 'saveDocument') {
        // R7 Desktop persists a locally-opened document through the editor's
        // own save entry point (editor.asc_Save -> CDocsCoApi.saveChanges ->
        // AscDesktopEditor.LocalFileSaveChanges). The plugin frame is
        // same-origin with the editor frame in desktop builds, so walk up the
        // frame chain and call it directly.
        try {
          var win = window
          for (var depth = 0; depth < 3 && win; depth++) {
            if (win.editor && typeof win.editor.asc_Save === 'function') {
              if (win.editor.asc_SetModified) win.editor.asc_SetModified(true)
              win.editor.asc_Save(false)
              sendReply(id, { success: true, via: 'asc_Save' })
              return
            }
            if (win === win.parent) break
            win = win.parent
          }
        } catch (err) {
          // Cross-origin frame chain: fall through to the documented method.
        }

        window.Asc.plugin.executeMethod('Save', [], function() {
          sendReply(id, { success: true, via: 'executeMethod' })
        })
        return
      }

      if (command === 'addParagraph') {
        Asc.scope.paraText = args.text || ''
        Asc.scope.paraStyle = args.style || 'Normal'
        window.Asc.plugin.callCommand(function() {
          var doc = Api.GetDocument()
          var p = Api.CreateParagraph()
          p.AddText(Asc.scope.paraText)
          if (Asc.scope.paraStyle !== 'Normal') {
            p.SetStyle(Asc.scope.paraStyle)
          }
          doc.Push(p)
        }, false, true, function(result) {
          sendReply(id, { success: true, result: result })
        })
        return
      }

      if (command === 'insertTable') {
        Asc.scope.rows = args.rows || 2
        Asc.scope.cols = args.cols || 2
        window.Asc.plugin.callCommand(function() {
          var doc = Api.GetDocument()
          var table = Api.CreateTable(Asc.scope.cols, Asc.scope.rows)
          doc.Push(table)
        }, false, true, function(result) {
          sendReply(id, { success: true, result: result })
        })
        return
      }

      if (command === 'getDocumentText') {
        window.Asc.plugin.callCommand(function() {
          var doc = Api.GetDocument()
          var text = ''
          var elementsCount = doc.GetElementsCount()
          for (var i = 0; i < elementsCount; i++) {
            var el = doc.GetElement(i)
            if (el.GetText) {
              text += el.GetText() + '\n'
            }
          }
          return text
        }, false, true, function(result) {
          sendReply(id, { success: true, text: result })
        })
        return
      }

      if (command === 'searchAndReplace') {
        Asc.scope.search = args.search || ''
        Asc.scope.replace = args.replace || ''
        window.Asc.plugin.callCommand(function() {
          var doc = Api.GetDocument()
          var elementsCount = doc.GetElementsCount()
          for (var i = 0; i < elementsCount; i++) {
            var el = doc.GetElement(i)
            if (el.SearchAndReplace) {
              el.SearchAndReplace(Asc.scope.search, Asc.scope.replace)
            }
          }
        }, false, true, function(result) {
          sendReply(id, { success: true, result: result })
        })
        return
      }

      sendReply(id, { success: false, error: 'Unsupported safeCommand: ' + command })
      return
    }

    if (action === 'callCommand') {
      const code = payload ? payload.code : ''
      try {
        const fn = new Function(code)
        window.Asc.plugin.callCommand(fn, false, true, function(result) {
          sendReply(id, { success: true, result: result })
        })
      } catch (err) {
        sendReply(id, { success: false, error: err.message })
      }
      return
    }

    sendReply(id, { success: false, error: 'Unknown action: ' + action })
  }

  function sendReply(id, data) {
    if (ws && ws.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify({ id: id, ...data }))
    }
  }

  window.Asc.plugin.button = function(id) {
    this.executeCommand('close', '')
  }
})(window, undefined)
