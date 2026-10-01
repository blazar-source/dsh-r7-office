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
  + 'и с открытым окном Р7 Desktop через MCP-инструменты (27 инструментов). '
  + 'ОБЩИЕ: '
  + 'r7_inspect — структура документа: заголовки, таблицы, листы, слайды, секции, колонтитулы, картинки, ссылки; '
  + 'r7_read — текст, markdown или диапазоны; при includeStyles=true возвращает нормализованное оформление, '
  + 'а не сырой OOXML, поэтому читайте его ПЕРЕД изменением; '
  + 'r7_create — создание НОВЫХ DOCX/XLSX/PPTX (существующий файл не перезаписывается без overwrite=true); '
  + 'r7_convert — нативная конвертация в PDF/HTML/TXT через движок x2t Р7; '
  + 'r7_validate — проверка целостности файла. '
  + 'DOCX: '
  + 'r7_edit — правка или удаление параграфа по индексу; '
  + 'r7_replace — поиск и замена текста со 100% сохранением стилей; '
  + 'r7_insert — вставка параграфов, заголовков, списков, разрывов; '
  + 'r7_table — таблицы: создание, правка, объединение ячеек, границы, заливка, ширины колонок; '
  + 'r7_docx_formatting — чтение оформления абзацев и ячеек: стиль, шрифт, размер, цвет, выравнивание, отступы, интервалы, списки; '
  + 'r7_docx_sections — размер страницы, ориентация, поля, колонки, разрывы страниц и секций; '
  + 'r7_docx_header_footer — чтение и правка колонтитулов (поля номеров страниц сохраняются); '
  + 'r7_docx_image — вставка PNG/JPEG/GIF с сохранением пропорций; '
  + 'r7_docx_hyperlink — чтение, вставка, переименование и удаление гиперссылок. '
  + 'XLSX: '
  + 'r7_sheet_read / r7_sheet_write — ячейки, диапазоны, формулы; лист адресуется по sheetName или sheetIndex; '
  + 'при date=true записывается РЕАЛЬНАЯ дата (serial), а не строка; '
  + 'r7_sheet_format — оформление ячейки или диапазона: шрифт, заливка, границы, выравнивание, перенос текста, '
  + 'числовые форматы (integer, decimal, currency, percent, date, datetime, custom), объединение, ширина колонок, высота строк; '
  + 'r7_sheet_add — новый лист в существующую книгу; '
  + 'r7_sheet_formula — вставка или обновление формулы. '
  + 'PPTX: '
  + 'r7_slide_read — нормализованная структура слайда: id, тип, геометрия (x/y/ширина/высота/поворот/z-order), '
  + 'текст, шрифт, заливка, обводка, выравнивание, абзацы; наследование от layout и master разрешается; '
  + 'r7_slide_create — новая презентация или ДОБАВЛЕНИЕ слайда на существующем layout; '
  + 'r7_slide_format — оформление и перемещение объекта на слайде: шрифт, геометрия, заливка, границы, списки, интервалы, текст; '
  + 'r7_slide_edit — дублирование, перемещение, переупорядочивание и удаление слайдов; '
  + 'r7_slide_object — добавление фигур, текстовых блоков и картинок, удаление и замена изображения. '
  + 'Р7 DESKTOP: '
  + 'r7_desktop_status / r7_desktop_selection / r7_desktop_exec — работа с открытым документом; '
  + 'r7_desktop_exec по умолчанию работает только с безопасным allowlist команд (safeCommand), '
  + 'произвольный DocScript/JS требует явного developerMode в конфигурации плагина. '
  + 'Порядок работы с файлами: r7_inspect → r7_read (при необходимости includeStyles) → правка → r7_validate → '
  + '(опционально) r7_convert в PDF. Исходный файл не перезаписывается без явного указания.'

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

    if (announce && prompt != null && typeof prompt.section === 'function') {
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
