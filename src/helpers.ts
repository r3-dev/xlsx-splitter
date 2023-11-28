export function exit(message: string): void {
  console.log(message)
  process.exit(1)
}

export function normalizeNumber(num: number): number {
  return Math.floor(Math.abs(num))
}
