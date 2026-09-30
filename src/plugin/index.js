/**
 * DeepSeek Harness Host Plugin for R7-Office (Р7-Офис).
 * Mounts r7_* tools on ctx.tools and injects agent guidance for processing
 * DOCX, XLSX, PPTX, PDF documents and connecting to live R7 Desktop.
 */

import { buildR7Tools } from '../mcp/tools.js'
import { DesktopBridge } from '../mcp/desktop-bridge.js'

/** Stable plugin name in DSH profile roster. */
export const name = 'r7-office'

/** Required service dependencies. */
export const inject = ['tools']

/** Guidance text injected into Agent system prompt. */
export const R7_GUIDANCE =
  'Установлен плагин dsh-r7-office: полноценная работа с документами Р7-Офис (DOCX, XLSX, PPTX, PDF) '
  + 'и открытым окном Р7 Desktop через MCP-инструменты. '
  + 'Инструменты: '
  + 'r7_inspect — анализ структуры, заголовков, таблиц, листов и слайдов; '
  + 'r7_read — чтение структурированного текста, markdown или диапазонов данных; '
  + 'r7_create — создание новых DOCX/XLSX/PPTX по нативным шаблонам Р7; '
  + 'r7_edit — точечное редактирование или удаление параграфов; '
  + 'r7_replace — поиск и замена текста со 100% сохранением стилей и форматирования; '
  + 'r7_insert — вставка параграфов, заголовков, списков, разрывов страниц; '
  + 'r7_table — создание, чтение, правка таблиц и ячеек; '
  + 'r7_sheet_read / r7_sheet_write / r7_sheet_formula — работа с ячейками, диапазонами и формулами XLSX; '
  + 'r7_slide_create / r7_slide_edit — управление слайдами PPTX; '
  + 'r7_convert — нативная конвертация в PDF/HTML/TXT через движок x2t Р7; '
  + 'r7_validate — проверка целостности файлов; '
  + 'r7_desktop_status / r7_desktop_selection / r7_desktop_exec — управление открытым документом в Р7 Desktop. '
  + 'r7_desktop_exec по умолчанию работает только с безопасным allowlist команд (safeCommand); произвольный DocScript/JS '
  + 'требует явного developerMode в конфигурации плагина. '
  + 'Порядок работы с файлами: r7_inspect → r7_read → r7_edit/r7_replace/r7_insert/r7_table → r7_validate → (опционально) r7_convert в PDF.'

/**
 * Mount the R7 plugin into DeepSeek Harness context.
 * @param ctx - Harness context
 * @param config - Plugin configuration options
 */
export function apply(ctx, config = {}) {
  const desktopBridge = new DesktopBridge({
    port: config.desktopBridgePort || 7888,
    // Arbitrary DocScript execution is OFF unless the operator opts in.
    developerMode: config.developerMode === true,
    r7Path: config.r7Path
  })
  const announce = config.announceToAgent !== false
  const prompt = ctx.get?.('systemPrompt')

  const tools = buildR7Tools({
    r7Path: config.r7Path,
    desktopBridge
  })

  const mount = () => {
    const disposers = tools.map((tool) => ctx.tools.register(tool))

    if (announce && prompt !== undefined && typeof prompt.section === 'function') {
      disposers.push(
        prompt.section({
          name: 'plugin:dsh-r7-office',
          order: 140,
          text: R7_GUIDANCE
        })
      )
    }

    // Auto-start desktop bridge if enabled
    if (config.enableDesktopBridge !== false) {
      desktopBridge.start().catch((err) => {
        // Silently handle port bind fallback
      })
    }

    // One concise activation line: it is the only positive evidence a boot log
    // can carry that the tools really reached the Harness tool registry.
    if (config.quiet !== true) {
      const names = tools.map((tool) => tool.name).join(', ')
      console.log(
        `[r7-office] зарегистрировано инструментов: ${tools.length} (${names}); `
        + `desktop bridge port=${desktopBridge.port}, developerMode=${desktopBridge.developerMode}`
      )
    }

    return () => {
      desktopBridge.stop().catch(() => {})
      for (const dispose of disposers) {
        try { dispose() } catch {}
      }
    }
  }

  if (typeof ctx.effect === 'function') {
    ctx.effect(mount, 'dsh-r7-office: surfaces')
    return
  }
  mount()
}
