import { NextRequest } from 'next/server';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { promises as fs } from 'node:fs';
import path from 'node:path';

export const runtime = 'nodejs';

type Payload = {
  teacher_name: string;
  subject: string;
  grade: string;
  week: string;
  lesson_title: string;
  objectives: string[];
  materials: string[];
  steps: string[];
  assessment: string;
  homework: string;
};

function safeArray(value: unknown): string[] {
  return Array.isArray(value) ? value.map((v) => String(v)) : [];
}

export async function POST(req: NextRequest) {
  try {
    const body = (await req.json()) as Partial<Payload>;
    const data: Payload = {
      teacher_name: String(body.teacher_name || ''),
      subject: String(body.subject || ''),
      grade: String(body.grade || ''),
      week: String(body.week || ''),
      lesson_title: String(body.lesson_title || ''),
      objectives: safeArray(body.objectives),
      materials: safeArray(body.materials),
      steps: safeArray(body.steps),
      assessment: String(body.assessment || ''),
      homework: String(body.homework || ''),
    };

    // Basic validation
    if (!data.teacher_name || !data.subject || !data.grade || !data.week || !data.lesson_title) {
      return new Response('Missing required fields', { status: 400 });
    }

    const templatePath = path.join(process.cwd(), 'templates', 'ministry_template.docx');
    const templateBuffer = await fs.readFile(templatePath);

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

    const templateData = {
      teacher_name: data.teacher_name,
      subject: data.subject,
      grade: data.grade,
      week: data.week,
      lesson_title: data.lesson_title,
      objectives: data.objectives,
      materials: data.materials,
      steps: data.steps,
      assessment: data.assessment,
      homework: data.homework,
      // Convenience derived values
      today: new Date().toLocaleDateString('ar-EG'),
    };

    doc.setData(templateData);
    try {
      doc.render();
    } catch (error: any) {
      console.error('Docxtemplater render error', error);
      return new Response('Template rendering failed', { status: 500 });
    }

    const out = doc.getZip().generate({ type: 'arraybuffer' });

    const filename = `خطة_${data.subject}_${data.lesson_title}.docx`;
    return new Response(out, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
        'Cache-Control': 'no-store',
      },
    });
  } catch (err) {
    console.error(err);
    return new Response('Invalid request', { status: 400 });
  }
}
