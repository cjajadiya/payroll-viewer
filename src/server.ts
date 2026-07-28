import express, { Request, Response, NextFunction } from 'express';
import multer from 'multer';
import * as xlsx from 'xlsx';
import { z } from 'zod';
import cors from 'cors';
import { PrismaClient } from '@prisma/client';

const app = express();
const prisma = new PrismaClient();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

app.use(cors());
app.use(express.json());

const ShiftSchema = z.object({
  externalEmpId: z.string().min(1),
  clockIn: z.string(),
  clockOut: z.string(),
  regularHours: z.number().nonnegative(),
  overtimeHours: z.number().nonnegative(),
});

type ShiftInput = z.infer<typeof ShiftSchema>;

app.post('/api/v1/payroll/upload', upload.single('file'), async (req: Request, res: Response, next: NextFunction): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'Missing file upload.' });
      return;
    }

    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[][] = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const validShifts: ShiftInput[] = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0]) continue;

      const parsed = ShiftSchema.safeParse({
        externalEmpId: String(row[0]).trim(),
        clockIn: new Date(row[2]).toISOString(),
        clockOut: new Date(row[3]).toISOString(),
        regularHours: Number(row[6]) || 0,
        overtimeHours: Number(row[8]) || 0,
      });

      if (parsed.success) {
        validShifts.push(parsed.data);
      }
    }

    res.status(200).json({
      status: 'success',
      recordsParsed: validShifts.length,
      data: validShifts,
    });
  } catch (err) {
    next(err);
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`[Payroll API] Server running on http://localhost:${PORT}`);
});
