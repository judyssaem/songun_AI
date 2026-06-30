export interface FileObject {
  id: string;
  name: string;
  type: string;
  size: number;
  base64: string;
}

export interface ScoreSet {
  score: number;
  max: number;
  normalized: number | null; // 0-5 scaled score
  evidence: string;
}

export interface DomainScores {
  "지식이해": ScoreSet;
  "과정기능": ScoreSet;
  "가치태도": ScoreSet;
}

export interface AnalysisResult {
  detected_name: string;
  rubric_mapping: {
    "지식이해": string[];
    "과정기능": string[];
    "가치태도": string[];
  };
  first_scores: DomainScores;
  second_scores: DomainScores;
  growth_analysis: {
    "지식이해": string;
    "과정기능": string;
    "가치태도": string;
  };
  overall_summary: string;
  teaching_feedback: string;
  achievement_level: {
    first: "A" | "B" | "C";
    second: "A" | "B" | "C";
  };
}

export interface StudentData {
  id: string;
  name: string;
  className: string;
  number: string;
  firstFiles: FileObject[];
  secondFiles: FileObject[];
  analysis: null | {
    status: "running" | "done" | "error";
    result?: AnalysisResult;
    error?: string;
    analyzedAt: number;
  };
  createdAt: number;
}

export interface SettingData {
  id: "main";
  apiKey: string;
  model: string;
  rememberKey: boolean;
  rubric1Mode: "file" | "text";
  rubric1Text: string;
  rubric1Files: FileObject[];
  rubric2Mode: "file" | "text";
  rubric2Text: string;
  rubric2Files: FileObject[];
}
