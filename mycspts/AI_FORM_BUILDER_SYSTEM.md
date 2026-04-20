# 🤖 ระบบการนำเข้าเอกสารและสร้างแบบฟอร์มด้วย AI (AI-Powered Form Builder)

เอกสารนี้รวบรวมข้อมูลและซอร์สโค้ดที่เกี่ยวข้องกับฟีเจอร์ "การสร้างแบบฟอร์มด้วย AI จากไฟล์เอกสาร" ของระบบ EvalPro ซึ่งทำหน้าที่แปลงเอกสารเก่าเช่น ไฟล์ Word (.docx) หรือ PDF (.pdf) ให้กลายเป็นแบบฟอร์มออนไลน์ (Digital Form) ได้โดยอัตโนมัติ

---

## 1. 📍 ภาพรวมกระบวนการทำงาน (System Workflow)

ระบบนี้มีลำดับการทำงาน (Workflow) 3 ขั้นตอนหลัก ดังนี้:
1. **Frontend (User Interface) 🖥️**: แอดมินหรือผู้สร้างแบบประเมินทำการกดปุ่ม "ดึงข้อมูลจากไฟล์" แล้วเลือกไฟล์เอกสาร `.docx` หรือ `.pdf` ผ่านหน้า Dashboard
2. **Backend API (Document Parsing) ⚙️**: ฝั่งเซิร์ฟเวอร์จะรับไฟล์มา แล้วใช้ไลบรารีพิเศษ (`mammoth` และ `pdf-parse`) เพื่อแกะเอา "ข้อความดิบ (Raw Text)" ออกมาจากเอกสาร
3. **AI Logic (Data Structuring) 🧠**: นำข้อความดิบส่งไปให้ **Google Gemini AI** วิเคราะห์ และจัดโครงสร้างใหม่กลับมาเป็นรูปแบบ **JSON** ตามที่กำหนดไว้ (แยกประเภทคำถาม, ดึงตัวเลือก ฯลฯ) เพื่อให้ Frontend นำไปปริ้นต์แบบฟอร์มและบันทึกลง Database ได้ทันที

---

## 2. 🧠 โค้ดส่วนหลัก AI Logic (การสกัดข้อมูลเป็นแบบฟอร์ม)

**พิกัดไฟล์:** `src/lib/gemini.ts`

ไฟล์นี้เป็นหัวใจสำคัญที่ทำการสื่อสารกับ Google Gemini AI โดยรับข้อความเข้ามา และออกคำสั่ง (Prompt) ให้ AI สวมบทบาทเป็นผู้สร้างฟอร์ม แล้วคายข้อมูลออกมาให้ตรงเป๊ะตามโครงสร้าง JSON

```typescript
import { GoogleGenerativeAI } from "@google/generative-ai";

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || "");

export async function extractTemplateFromText(text: string) {
  // เลือกใช้โมเดลที่รวดเร็วและฉลาด (flash)
  const model = genAI.getGenerativeModel({ model: "gemini-flash-latest" });

  // สั่งงานด้วย Prompt เพื่อให้ AI จัดโครงสร้างให้ออกมาเป็น JSON ตามเงื่อนไขเป๊ะๆ
  const prompt = `
    You are an expert at designing evaluation forms. 
    Analyze the following raw text extracted from a document and convert it into a structured evaluation template JSON.
    
    The JSON structure MUST follow this format:
    {
      "title": "Title of the evaluation",
      "description": "Short description",
      "questions": [
        {
          "label": "Question text",
          "type": "text" | "textarea" | "radio" | "checkbox" | "dropdown" | "file" | "linear_scale" | "rating" | "date" | "time",
          "scale": 10, // required if type is rating or linear_scale (default 10)
          "options": ["Option 1", "Option 2"], // required if type is radio, checkbox, or dropdown
          "minLabel": "Start Label", // optional for linear_scale
          "maxLabel": "End Label" // optional for linear_scale
        }
      ]
    }

    Raw Text:
    """
    ${text}
    """

    Return ONLY the raw JSON string.
  `;

  // ส่งข้อมูลให้ AI และรอรับคำตอบ
  const result = await model.generateContent(prompt);
  const response = await result.response;
  
  // ตัดข้อความขยะ/แท็ก Markdown (เช่น ```json ... ```) ออกไปให้เหลือแต่ JSON เพียวๆ
  const jsonString = response.text().replace(/```json|```/g, "").trim();

  // ป้องกัน Error หาก AI ตอบไม่ตรง format
  try {
    return JSON.parse(jsonString);
  } catch (parseError: unknown) {
    console.error("Failed to parse AI response as JSON", jsonString, parseError);
    throw new Error("AI response was not a valid JSON");
  }
}
```

---

## 3. ⚙️ โค้ดส่วน Backend API (การรับไฟล์และแยกข้อความดิบ)

**พิกัดไฟล์:** `src/app/api/admin/form-templates/import/route.ts`

ตัวรับข้อมูล (Endpoint) จากฝั่ง Frontend ทำหน้าที่ตรวจสอบสิทธิ์ (Auth) จากนั้นพินิจพิเคราะห์ว่าไฟล์ที่ได้มาเป็น .docx หรือ .pdf เพื่อใช้เครื่องมือแกะข้อความที่ถูกต้อง หากมีส่วนไหนผิดพลาด AI จะไม่ได้รับการกระตุ้น เพื่อลดภาระระบบ

```typescript
import { NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import mammoth from 'mammoth';        // สำหรับแกะข้อความจาก .docx
import { PDFParse } from 'pdf-parse'; // สำหรับแกะข้อความจาก .pdf
import { extractTemplateFromText } from '@/lib/gemini'; // เรียกใช้ช้ฟังก์ชัน AI

export async function POST(req: Request) {
    try {
        // 1. ตรวจสอบสิทธิ์ ว่าเข้าใช้งานด้วยแอดมินเท่านั้น
        const session = await getServerSession(authOptions);
        if (!session || session.user.role !== 'admin') {
            return NextResponse.json({ message: 'Unauthorized' }, { status: 401 });
        }

        // 2. รับไฟล์ที่อัปโหลดมา (formData)
        const formData = await req.formData();
        const file = formData.get('file') as File;

        if (!file) {
            return NextResponse.json({ message: 'No file uploaded' }, { status: 400 });
        }

        const buffer = Buffer.from(await file.arrayBuffer());
        let text = '';

        // 3. ตรวจสอบนามสกุลไฟล์ และแกะข้อความด้วยเครื่องมือที่เหมาะสม
        if (file.name.endsWith('.docx')) {
            const result = await mammoth.extractRawText({ buffer });
            text = result.value;
        } else if (file.name.endsWith('.pdf')) {
            const parser = new PDFParse({ data: buffer });
            const result = await parser.getText();
            text = result.text;
        } else {
            return NextResponse.json({ message: 'Unsupported file format. Please upload .docx or .pdf' }, { status: 400 });
        }

        // หากเอกสารว่างเปล่า
        if (!text.trim()) {
            return NextResponse.json({ message: 'Document is empty or could not be read' }, { status: 400 });
        }

        if (!process.env.GEMINI_API_KEY) {
            return NextResponse.json({ message: 'Gemini API key is not configured' }, { status: 500 });
        }

        // 4. ส่งข้อความดิบไปให้ Gemini AI และรับกลับมาเป็นชุด Form Template
        const template = await extractTemplateFromText(text);

        // 5. ส่งหน้าตา Form กลับไปให้ Frontend ทำการพรีวิว
        return NextResponse.json(template);
    } catch (error: unknown) {
        console.error('Import Error:', error);
        const errorMessage = error instanceof Error ? error.message : 'Internal Server Error';
        return NextResponse.json({ message: errorMessage }, { status: 500 });
    }
}
```

---

## 4. 🖥️ โค้ดส่วน Frontend (การกดใช้งาน AI จากหน้าเว็บ)

**พิกัดไฟล์ (ตัวอย่างเฉพาะฟังก์ชันที่เกี่ยวข้อง):** `src/app/admin/forms/page.tsx`

ฟังก์ชันหลักในการยิงไฟล์ไปให้ Backend พร้อมมี State การแสดงผล Loading (กำลังหมุนเพื่อรอ AI คิด) แจ้งเตือนผู้ใช้งานเมื่อสำเร็จ (ใช้ SweetAlert2)

```typescript
// ฟังก์ชันเมื่อผู้ใช้กดเลือกไฟล์ (input file change event)
const handleFileImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; // เอาไฟล์ช่องแรก
    if (!file) return;

    // สร้าง Form Data เตรียมยิงขึ้น API
    const formData = new FormData();
    formData.append('file', file);

    setImporting(true); // หมุน Loading AI กำลังประมวลผลจร้า
    try {
        // ยิง API
        const res = await axios.post('/api/admin/form-templates/import', formData);
        
        // เมื่อสำเร็จ ให้นำโครงสร้างจาก AI มาแทนที่ใน State ทันที พร้อมแสดง Alert!
        setTemplateData({ ...res.data, _id: templateData._id });
        Swal.fire({
            icon: 'success',
            title: 'วิเคราะห์ข้อมูลสำเร็จ',
            text: 'ระบบได้ดึงข้อมูลหัวข้อจากไฟล์ของคุณเรียบร้อยแล้ว กรุณาตรวจสอบปความถูกต้องก่อนบันทึก',
            confirmButtonColor: '#4361ee',
        });
    } catch (err: unknown) {
        let errorMessage = 'เกิดข้อผิดพลาดในการวิเคราะห์ไฟล์';
        if (axios.isAxiosError(err)) {
            errorMessage = err.response?.data?.message || errorMessage;
        }
        Swal.fire({
            icon: 'error',
            title: 'ไม่สามารถอ่านไฟล์ได้',
            text: errorMessage,
            confirmButtonColor: '#4361ee',
        });
    } finally {
        setImporting(false); // ปิด Loading
        e.target.value = ''; // เคลียร์ช่องให้เลือกไฟล์ใหม่ได้
    }
};

// ... โค้ดปุ่ม UI ตัวเลือกไฟล์ ...
<input
    type="file"
    id="importFile"
    className="d-none"
    accept=".docx,.pdf"
    onChange={handleFileImport}
/>
<label
    htmlFor="importFile"
    className={\`btn btn-white border shadow-sm rounded-pill px-4 py-2 small fw-bold text-slate-700 hover-up transition-all \${importing ? 'disabled' : ''}\`}
>
    {importing ? (
        <><span className="spinner-border spinner-border-sm me-2 text-primary"></span>กำลังวิเคราะห์ข้อมูล...</>
    ) : (
        <><i className="bi bi-cloud-arrow-up me-2 text-primary"></i>ดึงข้อมูลจากไฟล์ (Word/PDF)</>
    )}
</label>
```

---

## 5. 💡 หมวดฟีเจอร์ที่น่าสนใจในอนาคต (Future Interesting Features)

นอกจากฟีเจอร์ปัจจุบันแล้ว ระบบยังมีแนวคิดในการต่อยอดที่น่าสนใจและสามารถนำมาประยุกต์ใช้งานร่วมกับระบบฟอร์มนี้ได้อย่างมีประสิทธิภาพ ดังนี้:

### 📄 Smart Document Generation (ระบบสร้างเอกสารกลับอัตโนมัติ - Reverse AI Form)
**แนวคิด:** ต่อจากระบบแบบฟอร์มและการลงข้อมูลดิจิทัล บางหน่วยงานยังคงมีความจำเป็นที่จะต้องรวบรวม "ข้อมูลดิจิทัล" ที่ได้รับ กลับไปจัดวางบน "ใบเปเปอร์รูปแบบเดิม" (เช่น วท.1 หรือเอกสารคำร้อง) เพื่อนำไปเป็นเอกสารประกอบการส่งเซ็นอนุมัติตามระเบียบเดิม
**การทำงาน:** 
- เมื่อผู้ใช้งานกรอก Digital Form ผ่านระบบเสร็จสิ้น 
- ระบบจะนำคำตอบจากช่องต่างๆ (Variables) ไปหยอดใส่ในไฟล์ Template ต้นฉบับ (Word/PDF) แบบอัตโนมัติ 
- ระบบจะทำการ Generate ออกมาเป็นไฟล์เอกสารที่มีข้อมูลกรอกเรียบร้อยครบถ้วน พร้อมปรินต์หรือนำไปให้ที่ปรึกษาเซ็นอนุมัติต่อได้เลยทันที ช่วยอุดช่องโหว่ลดความยุ่งยากให้ผู้ใช้งาน

---

## 📌 สรุป
กลไกนี้ทำให้ระบบ EvalPro มีความทันสมัยและลดงานผู้ดูแลระบบไปได้อย่างมหาศาล ระบบเชื่อมต่อระหว่าง **Next.js API** + **File Parser (`mammoth` / `pdf-parse`)** และมีพระเอกเก่งๆอย่าง **Google Gemini AI (`gemini-flash-latest`)** ในการทุ่นแรงตีความข้อความจากเอกสารเดิมมาเป็นโครงสร้าง Database (`JSON`) ที่พร้อมให้แอปพลิเคชันใช้งานต่อทันทีครับ
