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
