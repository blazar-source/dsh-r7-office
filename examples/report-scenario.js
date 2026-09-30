/**
 * Example demonstrating the complete user scenario:
 * "Take Report.docx, inspect it, update Section 3 preserving formatting, add a table, validate, and convert to PDF."
 */

import path from 'node:path'
import os from 'node:os'
import fs from 'node:fs'
import { DocxEngine, R7Adapter } from '../src/r7/index.js'

async function runScenario() {
  const tmpDir = path.join(os.tmpdir(), 'dsh_r7_example')
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true })

  const originalDoc = path.join(tmpDir, 'Отчет.docx')
  const updatedDoc = path.join(tmpDir, 'Отчет_v2.docx')
  const pdfDoc = path.join(tmpDir, 'Отчет_v2.pdf')

  const docx = new DocxEngine()
  const adapter = new R7Adapter()

  console.log('1. Creating original document...')
  await docx.create(originalDoc, {
    title: 'Квартальный отчет компании',
    paragraphs: [
      { text: 'Раздел 1. Введение', style: 'Heading1', bold: true },
      'В первом квартале 2026 года компания показала стабильный рост.',
      { text: 'Раздел 2. Финансовые результаты', style: 'Heading1', bold: true },
      'Операционная прибыль составила 45.2 млн рублей.',
      { text: 'Раздел 3. Планы развития', style: 'Heading1', bold: true },
      'Черновой вариант планов развития: требует согласования и обновления.'
    ]
  })
  console.log('   Created:', originalDoc)

  console.log('\n2. Inspecting document structure...')
  const inspection = await docx.inspect(originalDoc)
  console.log(`   Paragraphs: ${inspection.paragraphsCount}, Headings: ${inspection.headingsCount}`)

  console.log('\n3. Reading document content...')
  const readRes = await docx.read(originalDoc, { format: 'markdown' })
  console.log('   Current content:\n', readRes.content)

  console.log('\n4. Updating Section 3 while preserving formatting...')
  await docx.replaceText(
    originalDoc,
    'Черновой вариант планов развития: требует согласования и обновления.',
    'Утвержден план расширения в 12 новых регионов и автоматизации документооборота.',
    { outputPath: updatedDoc }
  )

  console.log('\n5. Adding summary table...')
  await docx.table(updatedDoc, {
    action: 'create',
    rows: [
      ['Направление', 'Срок', 'Ответственный'],
      ['Внедрение Р7-Офис', 'Q2 2026', 'IT Департамент'],
      ['Интеграция ИИ агентов', 'Q3 2026', 'Команда DSH']
    ],
    outputPath: updatedDoc
  })

  console.log('\n6. Validating document integrity...')
  const val = await docx.validate(updatedDoc)
  console.log('   Validation result:', val.valid ? 'VALID (No errors)' : val.errors)

  console.log('\n7. Converting to PDF via R7 x2t...')
  const r7Info = await adapter.detect()
  if (r7Info.installed) {
    const conv = await adapter.convert(updatedDoc, pdfDoc)
    console.log(`   PDF successfully generated (${conv.timeMs}ms): ${pdfDoc}`)
  } else {
    console.log('   R7-Office not detected on host, skipping PDF conversion step.')
  }

  console.log('\nScenario finished successfully!')
}

runScenario().catch(console.error)
