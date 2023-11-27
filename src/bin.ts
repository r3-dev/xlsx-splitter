#!/usr/bin/env node
import fs, { constants } from 'node:fs/promises'
import path from 'node:path'
import xlsx from 'node-xlsx'
import yargs from 'yargs'
import { hideBin } from 'yargs/helpers'

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
    description: 'max number of rows to split'
  })
  .check(async (args) => {
    if (Number.isNaN(args.rows)) {
      throw new Error('Argument "--rows" must be a number.')
    }

    try {
      const stat = await fs.stat(args.file)

      if (!stat.isFile()) {
        throw new Error('Argument "--file" is directory. Expected .xlsx file.')
      }
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') {
        throw new Error(`File not found: ${args.file}`)
      }

      throw err
    }

    try {
      await fs.access(args.output, constants.W_OK)
    } catch {
      throw new Error(`Output directory is not writable: ${args.output}`)
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
  for (const row of excelPage.data) {
    if (!row || !row.length) continue

    if (!tableHead.length) {
      tableHead.push(...row)
    } else {
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
