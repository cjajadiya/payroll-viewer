import { NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

export async function POST(req: Request) {
  try {
    const data = await req.formData();
    const file: File | null = data.get('file') as unknown as File;

    if (!file) {
      return NextResponse.json({ error: "No file provided." }, { status: 400 });
    }

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    
    const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
    const parsedRowCount = rows.length;

    return NextResponse.json({ 
      status: "success", 
      message: "File parsed successfully.",
      rowsProcessed: parsedRowCount 
    }, { status: 200 });

  } catch (error) {
    console.error(error);
    return NextResponse.json({ error: "Internal Server Error during file processing." }, { status: 500 });
  }
}
