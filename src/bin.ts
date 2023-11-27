#!/usr/bin/env node
import fs from 'node:fs/promises'
import path from 'node:path'
import xlsx from 'node-xlsx'
import yargs from 'yargs'
import { hideBin } from 'yargs/helpers'

const argv = await yargs(hideBin(process.argv))
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
  .parse()

let fileCount = 0
let tableHead: string[] = []
let tableBody: string[][] = []

const xlsxFile = await fs.readFile(path.resolve(argv.file))
const workSheetsFromBuffer = xlsx.parse(xlsxFile)

for (const worksheet of workSheetsFromBuffer) {
  for (const worksheetKey in worksheet.data) {
    const data = worksheet.data[worksheetKey]
    if (!data || !data.length) continue

    if (worksheetKey === '0') {
      tableHead.push(...data)
    } else {
      tableBody.push(data)
    }

    if (tableBody.length === argv.rows) {
      await writeFile()
    }
  }
}

if (tableBody.length > 0) {
  await writeFile()
}

async function writeFile() {
  const buffer = xlsx.build([
    {
      name: path.basename(argv.file),
      data: [tableHead, ...tableBody],
      options: {}
    }
  ])

  tableBody = []
  const filePath = path.resolve(
    argv.output,
    path.basename(argv.file, '.xlsx') + `-${++fileCount}.xlsx`
  )
  await fs.writeFile(filePath, buffer)
}
