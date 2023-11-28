#!/usr/bin/env node
import fs, { constants } from 'node:fs/promises'
import path from 'node:path'
import xlsx from 'node-xlsx'
import yargs from 'yargs'
import { hideBin } from 'yargs/helpers'

import { exit, normalizeNumber } from './helpers.js'

const args = await yargs(hideBin(process.argv))
  .option('file', {
    alias: 'f',
    type: 'string',
    demandOption: true,
    description: '.xlsx file to parse'
  })
  .option('output', {
    alias: 'o',
    type: 'string',
    description: 'output directory',
    default: process.cwd()
  })
  .option('rows', {
    alias: 'r',
    type: 'number',
    demandOption: true,
    description: 'max number of rows to split',
    coerce: normalizeNumber
  })
  .option('offset', {
    alias: 's',
    type: 'number',
    description: 'number of start rows to skip',
    default: 0,
    coerce: normalizeNumber
  })
  .check(async (args) => {
    if (Number.isNaN(args.rows)) {
      exit('Argument "--rows" must be a number.')
    }

    if (args.rows === 0) {
      exit('Argument "--rows" must be greater than 0.')
    }

    if (Number.isNaN(args.offset)) {
      exit('Argument "--offset" must be a number.')
    }

    if (args.offset > args.rows) {
      exit('Argument "--offset" must be greater than or equal to "--rows".')
    }

    try {
      const stat = await fs.stat(args.file)
      if (!stat.isFile()) {
        exit('Argument "--file" is directory. Expected .xlsx file.')
      }
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') {
        exit(`File not found: ${args.file}`)
      }

      exit((err as Error).message)
    }

    try {
      await fs.access(args.output, constants.W_OK)
    } catch {
      exit(`Output directory is not writable: ${args.output}`)
    }

    return true
  })
  .parse()

console.log(`Parsing file: ${args.file}`)
console.log(`Splitting rows: ${args.rows}\n`)

const fileName = path.basename(args.file, '.xlsx')
const excelFile = await fs.readFile(path.resolve(args.file))
const parsedExcelFile = xlsx.parse(excelFile)

let fileCount = 0
const tableHead: string[] = []
const tableBody: string[][] = []

for (const excelPage of parsedExcelFile) {
  const rows = excelPage.data
  for (const rowKey in rows) {
    const rowIndex = Number(rowKey)
    const row = rows[rowIndex]
    if (!row || row.length === 0) continue

    if (tableHead.length === 0 && args.offset === rowIndex) {
      tableHead.push(...row)
    } else if (tableHead.length > 0) {
      tableBody.push(row)
    }

    if (tableBody.length === args.rows) {
      await createExcelFile()
    }
  }
}

if (tableBody.length > 0) {
  await createExcelFile()
}

console.log(`\nOutput directory: ${path.resolve(args.output)}`)

async function createExcelFile(): Promise<void> {
  const data = xlsx.build([
    {
      name: fileName,
      data: [tableHead, ...tableBody],
      options: {}
    }
  ])

  // clear array
  tableBody.length = 0

  const name = fileName + `-${++fileCount}.xlsx`
  const outputFilePath = path.resolve(args.output, name)
  await fs.writeFile(outputFilePath, data)
  console.log(`File created: ${name}`)
}
