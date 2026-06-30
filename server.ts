import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

// Helper to call Gemini API with retry and fallback
async function generateContentWithRetryAndFallback(ai: any, model: string, parts: any[], schema: any, maxRetries = 3) {
  let attempt = 0;
  let currentModel = model;

  // Map any non-existent model strings to real public Gemini model names
  if (currentModel === "gemini-3.5-flash" || currentModel === "gemini-3.1-flash-lite") {
    currentModel = "gemini-2.5-flash";
  } else if (currentModel === "gemini-3.1-pro-preview") {
    currentModel = "gemini-2.5-pro";
  }

  while (true) {
    try {
      console.log(`[Gemini Request] Model: ${currentModel} (Attempt ${attempt + 1}/${maxRetries + 1})`);
      const response = await ai.models.generateContent({
        model: currentModel,
        contents: { parts: parts },
        config: {
          responseMimeType: "application/json",
          responseSchema: schema,
          temperature: 0.0,
        }
      });
      return response;
    } catch (error: any) {
      attempt++;
      
      const errorStr = typeof error === "object" ? JSON.stringify(error) : String(error);
      const isUnavailable = 
        error?.status === "UNAVAILABLE" || 
        error?.code === 503 || 
        error?.message?.includes("503") || 
        error?.message?.includes("demand") ||
        error?.message?.includes("temporary") ||
        error?.message?.includes("RESOURCE_EXHAUSTED") ||
        error?.status === "RESOURCE_EXHAUSTED" ||
        errorStr.includes("503") ||
        errorStr.includes("UNAVAILABLE") ||
        errorStr.includes("demand") ||
        errorStr.includes("RESOURCE_EXHAUSTED");

      if (isUnavailable) {
        if (attempt <= maxRetries) {
          const delay = Math.pow(2, attempt) * 1000 + Math.random() * 1000;
          console.warn(`[Gemini Retry] Model ${currentModel} experienced high demand/503. Retrying in ${delay.toFixed(0)}ms (Attempt ${attempt}/${maxRetries})... Error: ${error.message || errorStr}`);
          await new Promise((resolve) => setTimeout(resolve, delay));
          continue;
        }

        // Fallback flow if we exhausted all retries on the current model
        if (currentModel === "gemini-2.5-pro") {
          console.warn(`[Gemini Fallback] Pro model failed after retries. Switching to fallback model gemini-2.5-flash...`);
          currentModel = "gemini-2.5-flash";
          attempt = 0; // reset attempts for fallback model
          maxRetries = 2; // give 2 retries for fallback model
          await new Promise((resolve) => setTimeout(resolve, 1500));
          continue;
        } else if (currentModel === "gemini-2.5-flash" || currentModel === "gemini-2.0-flash") {
          console.warn(`[Gemini Fallback] Flash model failed after retries. Switching to highly available fallback model gemini-1.5-flash...`);
          currentModel = "gemini-1.5-flash";
          attempt = 0; // reset attempts for fallback model
          maxRetries = 2; // give 2 retries for fallback model
          await new Promise((resolve) => setTimeout(resolve, 1500));
          continue;
        }
      }
      
      // If not a retryable error or we exhausted all retries, throw the final error
      throw error;
    }
  }
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // JSON Body Parser with high limit for Base64 image payloads
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  // API Route for Gemini analysis
  app.post("/api/analyze", async (req, res) => {
    try {
      const { apiKey, model, rubric1, rubric2, student } = req.body;

      // 1. Resolve API Key (Client provided takes precedence, fallback to env)
      const effectiveApiKey = apiKey?.trim() || process.env.GEMINI_API_KEY;

      if (!effectiveApiKey) {
        return res.status(400).json({
          error: "Gemini API 인증 키가 설정되지 않았습니다. 관리자 설정에서 GEMINI_API_KEY 환경 변수를 등록해 주세요.",
        });
      }

      // 2. Initialize Google Gen AI SDK
      const ai = new GoogleGenAI({
        apiKey: effectiveApiKey,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build",
          },
        },
      });

      // 3. Prepare multimodal parts
      const parts: any[] = [];

      // A. Add 1st Rubric information
      parts.push({ text: "=== [1차 평가 루브릭] ===" });
      if (rubric1.mode === "text") {
        parts.push({ text: rubric1.text || "(루브릭 텍스트 없음)" });
      } else if (rubric1.files && rubric1.files.length > 0) {
        for (const f of rubric1.files) {
          parts.push({
            inlineData: {
              mimeType: f.type,
              data: f.base64,
            },
          });
          parts.push({ text: `(파일명: ${f.name})` });
        }
      }

      // B. Add 2nd Rubric information
      parts.push({ text: "=== [2차 평가 루브릭] ===" });
      if (rubric2.mode === "text") {
        parts.push({ text: rubric2.text || "(루브릭 텍스트 없음)" });
      } else if (rubric2.files && rubric2.files.length > 0) {
        for (const f of rubric2.files) {
          parts.push({
            inlineData: {
              mimeType: f.type,
              data: f.base64,
            },
          });
          parts.push({ text: `(파일명: ${f.name})` });
        }
      }

      // C. Add Student's 1st Round Papers
      parts.push({ text: `=== [학생 ${student.name}의 1차 평가지] ===` });
      if (student.firstFiles && student.firstFiles.length > 0) {
        for (const f of student.firstFiles) {
          parts.push({
            inlineData: {
              mimeType: f.type,
              data: f.base64,
            },
          });
          parts.push({ text: `(1차 평가지 파일명: ${f.name})` });
        }
      } else {
        parts.push({ text: "(1차 평가지 파일 없음)" });
      }

      // D. Add Student's 2nd Round Papers
      parts.push({ text: `=== [학생 ${student.name}의 2차 평가지] ===` });
      if (student.secondFiles && student.secondFiles.length > 0) {
        for (const f of student.secondFiles) {
          parts.push({
            inlineData: {
              mimeType: f.type,
              data: f.base64,
            },
          });
          parts.push({ text: `(2차 평가지 파일명: ${f.name})` });
        }
      } else {
        parts.push({ text: "(2차 평가지 파일 없음)" });
      }

      // E. Add Analysis Instruction
      const analysisInstruction = `당신은 초등학교 수학 서·논술형 평가를 분석하는 전문가입니다.

주어진 1차 평가 루브릭, 2차 평가 루브릭, 그리고 한 학생의 1차/2차 평가지를 분석하여,
다음 3개 영역으로 성장을 평가하세요.
- 지식·이해 (Knowledge & Understanding)
- 과정·기능 (Process & Skills)
- 가치·태도 (Values & Attitudes)

[중요: 일관성 및 정확성 극대화 지침]
- 같은 파일(평가지, 루브릭)에 대해 매번 분석할 때마다 완전히 일관되고 객관적인 점수와 분석 결과를 도출해야 합니다.
- 주관적인 해석을 최대한 지양하고, 제공된 1차/2차 루브릭의 채점 기준과 한계 점수 배점 표를 100% 엄격하고 균일하게 적용하여 정확하게 산출하세요.

[작업 1: 루브릭 매핑]
루브릭의 각 채점 요소를 위 3개 영역 중 어디에 속하는지 판단해 분류하세요.
예시 기준:
- "막대그래프 요소 이해하기" → 지식·이해
- "막대그래프 그리기" → 과정·기능
- "막대그래프 해석하기 (질문 만들고 답하기)" → 과정·기능
- "새롭게 알게 된 점 + 생활 속 활용 사례" → 가치·태도
하나의 채점 요소가 두 영역에 걸칠 수 있으면 더 비중이 큰 쪽으로 배정하세요.

[작업 2: 채점]
학생의 응답을 루브릭에 따라 점수화하세요.
- 손글씨와 그림을 가능한 한 정확히 판독합니다.
- 막대그래프는 가로축/세로축 라벨, 눈금 단위 선택, 막대 길이 정확성, 제목 모두 평가합니다.
- 채점 요소를 영역별로 합산하여 raw score를 산출합니다.
  예: 과정·기능에 "그리기(3점)"와 "해석하기(2점)"가 매핑되면 → max=5

[작업 3: 5점 척도 환산 ⭐중요]
서로 다른 영역의 만점(예: 지식이해 1점 vs 과정기능 5점)이 방사형 그래프에서
시각적 왜곡을 일으키지 않도록, 각 영역의 raw score를 0-5점 척도로 환산합니다.

공식: normalized = (raw_score / max_score) * 5
- 소수점 첫째 자리까지 반올림 (예: 4.7, 3.3, 5.0)
- max_score가 0인 경우 (해당 영역이 그 차시에 평가되지 않은 경우): normalized = null
- 1차에는 평가하지 않은 영역도 있을 수 있습니다. 그 경우 evidence에 "1차 루브릭에 해당 영역 평가 없음"을 명시하고 normalized를 null로 설정.

[작업 4: 성장 분석 및 긍정적 강점 평가 ⭐중요]
1차 → 2차의 변화를 영역별로 분석합니다. 단순히 점수가 올랐다는 표현은 피하고,
구체적인 학생 응답의 변화(예: "1차에서는 휴대전화라는 정답만 제시, 2차에서는 조사 → 표 → 그래프 → 의사결정의 전 과정을 수행")를 근거로 서술하세요.
- **긍정적 성장 리포트 관점 반영**: 1차 평가 대비 2차 평가에 대해서는 학생이 보여준 노력, 작은 발전, 세심한 기입, 풀이 과정에서의 시도, 발전 지향성 등 **긍정적인 측면을 최대한 많이 포착하여 칭찬과 지지하는 어조로 평가**하세요.
- 2차 평가지에서 보이는 학생의 작은 성취, 풀이의 정교함, 개선된 태도를 놓치지 않고 긍정적으로 서술해야 합니다.

[작업 5: 성취 수준 판정]
루브릭의 성취수준 기준(A/B/C)을 적용해 1차/2차 각각의 성취수준을 판정합니다.
기준이 명시되지 않은 경우 합계 점수 기준 70%↑ A, 40%↑ B, 그 이하 C.

[작업 6: 다음 단계 지도 제안]
이 학생의 강점과 약점을 바탕으로 다음 수업에서 활용할 수 있는 구체적 지도 제안을
2-3문장으로 작성하세요. 일반론이 아니라 이 학생의 응답에서 발견된 특징에 근거해야 합니다.

응답은 반드시 정해진 JSON 스키마만 출력하세요.`;

      parts.push({ text: `=== [분석 지시] ===\n${analysisInstruction}` });

      // 4. Define Response Schema
      const scoreSetSchema = {
        type: Type.OBJECT,
        properties: {
          score: { type: Type.NUMBER, description: "채점 결과 획득한 raw score" },
          max: { type: Type.NUMBER, description: "해당 영역의 만점 (raw score 기준)" },
          normalized: { 
            type: Type.NUMBER, 
            description: "0-5점 척도로 환산한 점수. 공식: (raw_score / max_score) * 5 (소수점 첫째짜리 반올림). 해당 영역 평가 대상이 아닐 경우 null" 
          },
          evidence: { type: Type.STRING, description: "채점 및 점수 환산의 구체적 근거 문구" }
        },
        required: ["score", "max", "evidence"]
      };

      const analysisSchema = {
        type: Type.OBJECT,
        properties: {
          detected_name: { type: Type.STRING, description: "평가지에서 인식된 학생 이름" },
          rubric_mapping: {
            type: Type.OBJECT,
            properties: {
              "지식이해": { type: Type.ARRAY, items: { type: Type.STRING }, description: "지식·이해 영역에 매핑된 루브릭 요소 리스트" },
              "과정기능": { type: Type.ARRAY, items: { type: Type.STRING }, description: "과정·기능 영역에 매핑된 루브릭 요소 리스트" },
              "가치태도": { type: Type.ARRAY, items: { type: Type.STRING }, description: "가치·태도 영역에 매핑된 루브릭 요소 리스트" }
            },
            required: ["지식이해", "과정기능", "가치태도"]
          },
          first_scores: {
            type: Type.OBJECT,
            properties: {
              "지식이해": scoreSetSchema,
              "과정기능": scoreSetSchema,
              "가치태도": scoreSetSchema
            },
            required: ["지식이해", "과정기능", "가치태도"]
          },
          second_scores: {
            type: Type.OBJECT,
            properties: {
              "지식이해": scoreSetSchema,
              "과정기능": scoreSetSchema,
              "가치태도": scoreSetSchema
            },
            required: ["지식이해", "과정기능", "가치태도"]
          },
          growth_analysis: {
            type: Type.OBJECT,
            properties: {
              "지식이해": { type: Type.STRING, description: "지식·이해 영역의 1차 대비 2차 성장 분석 내용 (구체적 변화 양상 근거 제시)" },
              "과정기능": { type: Type.STRING, description: "과정·기능 영역의 1차 대비 2차 성장 분석 내용 (구체적 변화 양상 근거 제시)" },
              "가치태도": { type: Type.STRING, description: "가치·태도 영역의 1차 대비 2차 성장 분석 내용 (구체적 변화 양상 근거 제시)" }
            },
            required: ["지식이해", "과정기능", "가치태도"]
          },
          overall_summary: { type: Type.STRING, description: "1차 대비 2차 성장 전반에 대한 요약 서평 (학생의 종합적인 성장 추이)" },
          teaching_feedback: { type: Type.STRING, description: "강약점에 기반한 다음 단계 학습 지도 제안 (구체적 실천 과제 2-3문장)" },
          achievement_level: {
            type: Type.OBJECT,
            properties: {
              first: { type: Type.STRING, description: "1차 성취수준 (A, B, C 중 선택)" },
              second: { type: Type.STRING, description: "2차 성취수준 (A, B, C 중 선택)" }
            },
            required: ["first", "second"]
          }
        },
        required: [
          "detected_name",
          "rubric_mapping",
          "first_scores",
          "second_scores",
          "growth_analysis",
          "overall_summary",
          "teaching_feedback",
          "achievement_level"
        ]
      };

      // 5. Call Gemini API
      // Supported models: gemini-2.5-flash is the default if unspecified or invalid
      const resolvedModel = model || "gemini-2.5-flash";

      const response = await generateContentWithRetryAndFallback(
        ai,
        resolvedModel,
        parts,
        analysisSchema,
        3
      );

      const responseText = response.text;
      if (!responseText) {
        throw new Error("Gemini API가 빈 응답을 반환했습니다.");
      }

      // 6. Parse and return JSON
      const resultJson = JSON.parse(responseText.trim());
      return res.json({ success: true, result: resultJson });

    } catch (error: any) {
      console.error("Gemini 분석 오류:", error);
      let userFriendlyError = error.message || "서버 분석 과정에서 알 수 없는 오류가 발생했습니다.";
      
      const errorStr = typeof error === "object" ? JSON.stringify(error) : String(error);
      if (
        error?.status === "UNAVAILABLE" || 
        error?.code === 503 || 
        error?.message?.includes("503") || 
        error?.message?.includes("demand") ||
        error?.message?.includes("temporary") ||
        error?.message?.includes("RESOURCE_EXHAUSTED") ||
        errorStr.includes("503") ||
        errorStr.includes("UNAVAILABLE") ||
        errorStr.includes("demand") ||
        errorStr.includes("RESOURCE_EXHAUSTED")
      ) {
        userFriendlyError = "현재 Google Gemini AI 모델의 사용량이 매우 많아 서비스가 지연되고 있습니다. 잠시 후 '재시도' 버튼을 누르시거나, [설정] 탭에서 다른 AI 모델(예: gemini-3.1-flash-lite)로 변경하여 진행해 주세요.";
      }

      return res.status(500).json({
        success: false,
        error: userFriendlyError
      });
    }
  });

  // Vite middleware for development, Static serve for production
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
