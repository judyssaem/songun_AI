import React, { useState, useEffect, useRef } from "react";
import { dbService } from "./lib/db";
import { FileObject, SettingData, StudentData, AnalysisResult } from "./types";
import * as XLSX from "xlsx";
import { 
  Settings, 
  Users, 
  BarChart3, 
  Upload, 
  Trash2, 
  Play, 
  Plus, 
  Loader2, 
  X, 
  Download, 
  RefreshCw, 
  Printer, 
  FileText, 
  ChevronRight, 
  AlertCircle, 
  CheckCircle2,
  Edit3
} from "lucide-react";

export default function App() {
  // --- States ---
  const [activeTab, setActiveTab] = useState<"settings" | "students" | "dashboard">("settings");
  const [settings, setSettings] = useState<SettingData>({
    id: "main",
    apiKey: "",
    model: "gemini-2.5-flash",
    rememberKey: false,
    rubric1Mode: "file",
    rubric1Text: "",
    rubric1Files: [],
    rubric2Mode: "file",
    rubric2Text: "",
    rubric2Files: [],
  });

  const [students, setStudents] = useState<StudentData[]>([]);
  const [selectedStudentId, setSelectedStudentId] = useState<string | null>(null);
  
  // Modals & Forms
  const [showAddModal, setShowAddModal] = useState(false);
  const [addMethod, setAddMethod] = useState<"single" | "excel">("single");
  const [showClearConfirmModal, setShowClearConfirmModal] = useState(false);
  const [studentToDelete, setStudentToDelete] = useState<StudentData | null>(null);
  const [newStudent, setNewStudent] = useState({ name: "", grade: "", classVal: "", number: "" });
  const [editingStudent, setEditingStudent] = useState<StudentData | null>(null);

  // Growth Dashboard edit states
  const [isEditingDashboard, setIsEditingDashboard] = useState(false);
  const [editDashboardData, setEditDashboardData] = useState<AnalysisResult | null>(null);

  useEffect(() => {
    setIsEditingDashboard(false);
    setEditDashboardData(null);
  }, [selectedStudentId]);

  // Statuses
  const [toasts, setToasts] = useState<Array<{ id: string; message: string; type: "success" | "error" | "info" }>>([]);
  const [apiProcessing, setApiProcessing] = useState<Record<string, boolean>>({});
  const [isAnalyzingAll, setIsAnalyzingAll] = useState(false);

  // File drag & drop states
  const [dragActiveId, setDragActiveId] = useState<string | null>(null);

  // --- Initial Data Loading ---
  useEffect(() => {
    async function loadData() {
      try {
        const loadedSettings = await dbService.getSettings();
        setSettings(loadedSettings);

        const loadedStudents = await dbService.getAllStudents();
        setStudents(loadedStudents);
        
        // Default select first student on dashboard if exists
        const completedStudent = loadedStudents.find(s => s.analysis?.status === "done");
        if (completedStudent) {
          setSelectedStudentId(completedStudent.id);
        } else if (loadedStudents.length > 0) {
          setSelectedStudentId(loadedStudents[0].id);
        }
      } catch (err) {
        showToast("데이터베이스를 읽어오는 중 오류가 발생했습니다.", "error");
      }
    }
    loadData();
  }, []);

  // --- Toast Handler ---
  const showToast = (message: string, type: "success" | "error" | "info" = "info") => {
    const id = Date.now().toString() + Math.random().toString(36).substr(2, 5);
    setToasts((prev) => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4000);
  };

  // --- Settings Actions ---
  const handleSettingsChange = (updates: Partial<SettingData>) => {
    const newSettings = { ...settings, ...updates };
    setSettings(newSettings);
    dbService.saveSettings(newSettings);
  };

  const handleApiKeyChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    handleSettingsChange({ apiKey: e.target.value });
  };

  const handleModelChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    handleSettingsChange({ model: e.target.value });
  };

  const handleRememberKeyChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    handleSettingsChange({ rememberKey: e.target.checked });
  };

  // Rubrics conversion & updates
  const handleRubricModeChange = (type: 1 | 2, mode: "file" | "text") => {
    if (type === 1) {
      handleSettingsChange({ rubric1Mode: mode });
    } else {
      handleSettingsChange({ rubric2Mode: mode });
    }
    showToast(`${type}차 루브릭 작성 모드가 변경되었습니다.`, "success");
  };

  const handleRubricTextChange = (type: 1 | 2, text: string) => {
    if (type === 1) {
      handleSettingsChange({ rubric1Text: text });
    } else {
      handleSettingsChange({ rubric2Text: text });
    }
  };

  const handleRubricFileUpload = async (type: 1 | 2, e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const uploadedFiles: FileObject[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const base64 = await fileToBase64(file);
        uploadedFiles.push({
          id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
          name: file.name,
          type: file.type,
          size: file.size,
          base64,
        });
      } catch (err) {
        showToast(`${file.name} 파일을 읽는 중 오류가 발생했습니다.`, "error");
      }
    }

    if (type === 1) {
      handleSettingsChange({ rubric1Files: [...settings.rubric1Files, ...uploadedFiles] });
    } else {
      handleSettingsChange({ rubric2Files: [...settings.rubric2Files, ...uploadedFiles] });
    }
    showToast(`${type}차 루브릭에 파일이 추가되었습니다.`, "success");
  };

  const handleDeleteRubricFile = (type: 1 | 2, fileId: string) => {
    if (type === 1) {
      const filtered = settings.rubric1Files.filter((f) => f.id !== fileId);
      handleSettingsChange({ rubric1Files: filtered });
    } else {
      const filtered = settings.rubric2Files.filter((f) => f.id !== fileId);
      handleSettingsChange({ rubric2Files: filtered });
    }
    showToast("루브릭 파일이 삭제되었습니다.", "info");
  };

  // --- Student Actions ---
  const handleAddStudentSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newStudent.name.trim()) {
      showToast("학생 이름을 입력해주세요.", "error");
      return;
    }

    const grade = newStudent.grade.trim();
    const classVal = newStudent.classVal.trim();
    let combinedClassName = "";
    if (grade && classVal) {
      const g = grade.endsWith("학년") ? grade : `${grade}학년`;
      const c = classVal.endsWith("반") ? classVal : `${classVal}반`;
      combinedClassName = `${g} ${c}`;
    } else if (grade) {
      combinedClassName = grade.endsWith("학년") ? grade : `${grade}학년`;
    } else if (classVal) {
      combinedClassName = classVal.endsWith("반") ? classVal : `${classVal}반`;
    } else {
      combinedClassName = "4반"; // Fallback default
    }

    const created: StudentData = {
      id: "std_" + Date.now().toString() + Math.random().toString(36).substr(2, 5),
      name: newStudent.name.trim(),
      className: combinedClassName,
      number: newStudent.number.trim(),
      firstFiles: [],
      secondFiles: [],
      analysis: null,
      createdAt: Date.now(),
    };

    try {
      await dbService.saveStudent(created);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      setNewStudent({ name: "", grade: "", classVal: "", number: "" });
      setShowAddModal(false);
      showToast(`${created.name} 학생이 추가되었습니다.`, "success");
      if (!selectedStudentId) {
        setSelectedStudentId(created.id);
      }
    } catch (err) {
      showToast("학생 추가에 실패했습니다.", "error");
    }
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        if (!bstr) {
          showToast("파일을 읽을 수 없습니다.", "error");
          return;
        }
        
        const workbook = XLSX.read(bstr, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // header: 1 options parsed to array of arrays
        const data = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
        if (data.length <= 1) {
          showToast("엑셀 파일에 유효한 학생 데이터가 없습니다. (최소 2줄 필요)", "error");
          return;
        }

        const headers = (data[0] || []).map((h: any) => String(h || "").trim());
        
        let gradeIdx = headers.findIndex((h: string) => h === "학년" || h.includes("학년"));
        let classIdx = headers.findIndex((h: string) => h === "반" || h.includes("반"));
        let numberIdx = headers.findIndex((h: string) => h === "번호" || h.includes("번호") || h.includes("학번"));
        let nameIdx = headers.findIndex((h: string) => h === "이름" || h.includes("이름") || h.includes("성명"));

        // Fallbacks if not found by name
        if (gradeIdx === -1) gradeIdx = 0;
        if (classIdx === -1) classIdx = 1;
        if (numberIdx === -1) numberIdx = 2;
        if (nameIdx === -1) nameIdx = 3;

        const parsedStudents: StudentData[] = [];
        const timestamp = Date.now();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const rawGrade = row[gradeIdx];
          const rawClass = row[classIdx];
          const rawNumber = row[numberIdx];
          const rawName = row[nameIdx];

          const nameStr = String(rawName !== undefined && rawName !== null ? rawName : "").trim();
          if (!nameStr) continue; // Skip empty name row

          const gradeStr = String(rawGrade !== undefined && rawGrade !== null ? rawGrade : "").trim();
          const classStr = String(rawClass !== undefined && rawClass !== null ? rawClass : "").trim();
          const numberStr = String(rawNumber !== undefined && rawNumber !== null ? rawNumber : "").trim();

          // Format class name beautifully
          let combinedClassName = "";
          if (gradeStr && classStr) {
            const g = gradeStr.endsWith("학년") ? gradeStr : `${gradeStr}학년`;
            const c = classStr.endsWith("반") ? classStr : `${classStr}반`;
            combinedClassName = `${g} ${c}`;
          } else if (gradeStr) {
            combinedClassName = gradeStr.endsWith("학년") ? gradeStr : `${gradeStr}학년`;
          } else if (classStr) {
            combinedClassName = classStr.endsWith("반") ? classStr : `${classStr}반`;
          } else {
            combinedClassName = "4반"; // Fallback default
          }

          const studentId = "std_" + timestamp + "_" + i + "_" + Math.random().toString(36).substr(2, 5);
          parsedStudents.push({
            id: studentId,
            name: nameStr,
            className: combinedClassName,
            number: numberStr,
            firstFiles: [],
            secondFiles: [],
            analysis: null,
            createdAt: timestamp + i,
          });
        }

        if (parsedStudents.length === 0) {
          showToast("유효한 학생 데이터가 없습니다. 열 이름을 확인해주세요. (학년, 반, 번호, 이름)", "error");
          return;
        }

        // Save all parsed students to database
        for (const std of parsedStudents) {
          await dbService.saveStudent(std);
        }

        const updatedList = await dbService.getAllStudents();
        setStudents(updatedList);
        
        // Auto select the first added student if nothing is selected
        if (!selectedStudentId && updatedList.length > 0) {
          setSelectedStudentId(parsedStudents[0].id);
        }

        showToast(`엑셀 파일에서 ${parsedStudents.length}명의 학생이 일괄 등록되었습니다.`, "success");
        setShowAddModal(false);
      } catch (err) {
        console.error("Excel import error:", err);
        showToast("엑셀 파일 처리 중 오류가 발생했습니다.", "error");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = ""; // reset input
  };

  const handleDeleteStudent = (id: string, name: string) => {
    const student = students.find(s => s.id === id);
    if (student) {
      setStudentToDelete(student);
    } else {
      setStudentToDelete({ id, name, className: "", number: "", firstFiles: [], secondFiles: [], analysis: null, createdAt: 0 });
    }
  };

  const confirmDeleteStudent = async () => {
    if (!studentToDelete) return;
    const { id, name } = studentToDelete;
    try {
      await dbService.deleteStudent(id);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      if (selectedStudentId === id) {
        const completed = updatedList.find(s => s.analysis?.status === "done");
        setSelectedStudentId(completed ? completed.id : (updatedList[0]?.id || null));
      }
      showToast(`${name} 학생 데이터가 삭제되었습니다.`, "info");
    } catch (err) {
      showToast("학생 데이터 삭제에 실패했습니다.", "error");
    } finally {
      setStudentToDelete(null);
    }
  };

  const handleEditStudentClick = (student: StudentData) => {
    setEditingStudent(student);
  };

  const handleEditStudentSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingStudent) return;
    if (!editingStudent.name.trim()) {
      showToast("이름은 필수 항목입니다.", "error");
      return;
    }

    try {
      await dbService.saveStudent(editingStudent);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      setEditingStudent(null);
      showToast("학생 정보가 수정되었습니다.", "success");
    } catch (err) {
      showToast("학생 정보 수정에 실패했습니다.", "error");
    }
  };

  // File uploads for specific student
  const handleStudentFileUpload = async (studentId: string, round: 1 | 2, e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const student = students.find((s) => s.id === studentId);
    if (!student) return;

    const uploadedFiles: FileObject[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const base64 = await fileToBase64(file);
        uploadedFiles.push({
          id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
          name: file.name,
          type: file.type,
          size: file.size,
          base64,
        });
      } catch (err) {
        showToast(`${file.name} 업로드 실패.`, "error");
      }
    }

    const updatedStudent = { ...student };
    if (round === 1) {
      updatedStudent.firstFiles = [...updatedStudent.firstFiles, ...uploadedFiles];
    } else {
      updatedStudent.secondFiles = [...updatedStudent.secondFiles, ...uploadedFiles];
    }

    // Reset analysis when files change, as it needs re-evaluation
    if (updatedStudent.analysis) {
      updatedStudent.analysis = null;
    }

    try {
      await dbService.saveStudent(updatedStudent);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      showToast(`${student.name} 학생의 ${round}차 평가지를 업로드했습니다.`, "success");
    } catch (err) {
      showToast("파일 저장을 완료하지 못했습니다.", "error");
    }
  };

  const handleDeleteStudentFile = async (studentId: string, round: 1 | 2, fileId: string) => {
    const student = students.find((s) => s.id === studentId);
    if (!student) return;

    const updatedStudent = { ...student };
    if (round === 1) {
      updatedStudent.firstFiles = updatedStudent.firstFiles.filter((f) => f.id !== fileId);
    } else {
      updatedStudent.secondFiles = updatedStudent.secondFiles.filter((f) => f.id !== fileId);
    }

    if (updatedStudent.analysis) {
      updatedStudent.analysis = null;
    }

    try {
      await dbService.saveStudent(updatedStudent);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      showToast("평가지 파일이 제거되었습니다.", "info");
    } catch (err) {
      showToast("파일 제거에 실패했습니다.", "error");
    }
  };

  // Drag & drop handlers
  const handleDragOver = (e: React.DragEvent, id: string) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActiveId(id);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActiveId(null);
  };

  const handleDrop = async (e: React.DragEvent, studentId: string, round: 1 | 2) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActiveId(null);

    const files = e.dataTransfer.files;
    if (!files || files.length === 0) return;

    const student = students.find((s) => s.id === studentId);
    if (!student) return;

    const uploadedFiles: FileObject[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const base64 = await fileToBase64(file);
        uploadedFiles.push({
          id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
          name: file.name,
          type: file.type,
          size: file.size,
          base64,
        });
      } catch (err) {
        showToast(`${file.name} 파일을 업로드하지 못했습니다.`, "error");
      }
    }

    const updatedStudent = { ...student };
    if (round === 1) {
      updatedStudent.firstFiles = [...updatedStudent.firstFiles, ...uploadedFiles];
    } else {
      updatedStudent.secondFiles = [...updatedStudent.secondFiles, ...uploadedFiles];
    }

    if (updatedStudent.analysis) {
      updatedStudent.analysis = null;
    }

    try {
      await dbService.saveStudent(updatedStudent);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      showToast(`${student.name} 학생의 ${round}차 평가지 ${uploadedFiles.length}개가 추가되었습니다.`, "success");
    } catch (err) {
      showToast("파일 업데이트 도중 저장 에러가 발생했습니다.", "error");
    }
  };

  // Helper base64 conversion with automatic image compression
  async function fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      // If the file is not an image (e.g. PDF), use regular FileReader without compression
      if (!file.type.startsWith("image/")) {
        const reader = new FileReader();
        reader.onload = () => {
          if (typeof reader.result === "string") {
            resolve(reader.result.split(",")[1]);
          } else {
            reject(new Error("Failed to convert file to base64"));
          }
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
        return;
      }

      // If it is an image, compress/resize it on an HTML Canvas to prevent "Request Entity Too Large"
      const img = new Image();
      img.src = URL.createObjectURL(file);
      img.onload = () => {
        URL.revokeObjectURL(img.src);
        let width = img.width;
        let height = img.height;
        const maxDim = 1000; // Limit max dimension to 1000px (retains excellent legibility while dramatically reducing size)

        if (width > maxDim || height > maxDim) {
          if (width > height) {
            height = Math.round((height * maxDim) / width);
            width = maxDim;
          } else {
            width = Math.round((width * maxDim) / height);
            height = maxDim;
          }
        }

        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;

        const ctx = canvas.getContext("2d");
        if (!ctx) {
          // Fallback to standard base64 if canvas is not supported
          const reader = new FileReader();
          reader.onload = () => {
            if (typeof reader.result === "string") {
              resolve(reader.result.split(",")[1]);
            } else {
              reject(new Error("Failed to convert file to base64"));
            }
          };
          reader.onerror = reject;
          reader.readAsDataURL(file);
          return;
        }

        // Fill background with white in case of transparent images (PNGs, etc.)
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);

        // Convert to optimized JPEG (60% quality is the industry standard sweet-spot for readability vs file size)
        const dataUrl = canvas.toDataURL("image/jpeg", 0.60);
        resolve(dataUrl.split(",")[1]);
      };
      img.onerror = () => {
        // Fallback to standard base64 on error
        const reader = new FileReader();
        reader.onload = () => {
          if (typeof reader.result === "string") {
            resolve(reader.result.split(",")[1]);
          } else {
            reject(new Error("Failed to convert file to base64"));
          }
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
      };
    });
  }

  // Helper function to compress existing base64 images on-the-fly before sending to API
  async function compressImageBase64(base64: string, mimeType: string, maxDim = 1000, quality = 0.60): Promise<string> {
    if (!mimeType || !mimeType.startsWith("image/")) {
      return base64;
    }
    // If the base64 is already very small (under 100KB), we can skip canvas drawing to make it even faster
    // 100KB in base64 is roughly 133,000 characters
    if (base64.length < 133000) {
      return base64;
    }
    return new Promise((resolve) => {
      const img = new Image();
      img.src = `data:${mimeType};base64,${base64}`;
      img.onload = () => {
        let width = img.width;
        let height = img.height;

        if (width > maxDim || height > maxDim) {
          if (width > height) {
            height = Math.round((height * maxDim) / width);
            width = maxDim;
          } else {
            width = Math.round((width * maxDim) / height);
            height = maxDim;
          }
        }

        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;

        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve(base64);
          return;
        }

        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);

        const dataUrl = canvas.toDataURL("image/jpeg", quality);
        resolve(dataUrl.split(",")[1]);
      };
      img.onerror = () => {
        resolve(base64);
      };
    });
  }

  // --- Growth Dashboard Edit Handlers ---
  const handleScoreChange = (
    round: "first_scores" | "second_scores",
    domain: "지식이해" | "과정기능" | "가치태도",
    field: "score" | "max" | "normalized" | "evidence",
    value: any
  ) => {
    setEditDashboardData((prev) => {
      if (!prev) return null;
      const currentRoundScores = prev[round] || {};
      const currentDomainScores = currentRoundScores[domain] || { score: 0, max: 5, normalized: 0, evidence: "" };
      
      let updatedVal = value;
      if (field === "score" || field === "max" || field === "normalized") {
        updatedVal = value === "" ? "" : Number(value);
      }

      const updatedDomain = {
        ...currentDomainScores,
        [field]: updatedVal
      };

      // Auto-recalculate normalized if score or max changed
      if (field === "score" || field === "max") {
        const scoreNum = field === "score" ? Number(value) : Number(currentDomainScores.score);
        const maxNum = field === "max" ? Number(value) : Number(currentDomainScores.max);
        if (maxNum > 0) {
          updatedDomain.normalized = Number(((scoreNum / maxNum) * 5).toFixed(1));
        }
      }

      return {
        ...prev,
        [round]: {
          ...currentRoundScores,
          [domain]: updatedDomain
        }
      };
    });
  };

  const handleSaveDashboardData = async () => {
    if (!selectedStudent || !editDashboardData) return;
    try {
      const updatedStudent = {
        ...selectedStudent,
        analysis: {
          ...selectedStudent.analysis!,
          result: editDashboardData,
          analyzedAt: Date.now()
        }
      };
      await dbService.saveStudent(updatedStudent);
      const updatedList = await dbService.getAllStudents();
      setStudents(updatedList);
      setIsEditingDashboard(false);
      setEditDashboardData(null);
      showToast(`${selectedStudent.name} 학생의 분석 결과가 저장되었습니다.`, "success");
    } catch (err) {
      showToast("분석 결과 저장에 실패했습니다.", "error");
    }
  };

  // --- API / AI Analysis Engine ---
  const analyzeSingleStudent = async (studentId: string, quiet: boolean = false): Promise<boolean> => {
    const student = students.find((s) => s.id === studentId);
    if (!student) return false;

    // Check pre-requisites
    if (student.firstFiles.length === 0 || student.secondFiles.length === 0) {
      if (!quiet) {
        showToast(`${student.name} 학생의 1차 및 2차 평가지를 먼저 모두 업로드해 주세요.`, "error");
      }
      return false;
    }

    const hasRubric1 = settings.rubric1Mode === "text" ? settings.rubric1Text.trim() : settings.rubric1Files.length > 0;
    const hasRubric2 = settings.rubric2Mode === "text" ? settings.rubric2Text.trim() : settings.rubric2Files.length > 0;

    if (!hasRubric1 || !hasRubric2) {
      if (!quiet) {
        showToast("설정 탭에서 1차 및 2차 평가 루브릭을 먼저 작성하거나 업로드해 주세요.", "error");
      }
      return false;
    }

    // Mark as running in UI
    setApiProcessing((prev) => ({ ...prev, [studentId]: true }));
    
    // Update local state and DB status to 'running'
    const updatedWithRunning = {
      ...student,
      analysis: {
        status: "running" as const,
        analyzedAt: Date.now(),
      }
    };
    
    // Optimistically update lists
    setStudents(prev => prev.map(s => s.id === studentId ? updatedWithRunning : s));
    await dbService.saveStudent(updatedWithRunning);

    try {
      // Compress rubric & student files on-the-fly before API request to avoid Vercel payload limit (4.5MB)
      const compressedRubric1Files = await Promise.all(
        (settings.rubric1Files || []).map(async (f) => ({
          ...f,
          base64: await compressImageBase64(f.base64, f.type),
        }))
      );

      const compressedRubric2Files = await Promise.all(
        (settings.rubric2Files || []).map(async (f) => ({
          ...f,
          base64: await compressImageBase64(f.base64, f.type),
        }))
      );

      const compressedFirstFiles = await Promise.all(
        (student.firstFiles || []).map(async (f) => ({
          ...f,
          base64: await compressImageBase64(f.base64, f.type),
        }))
      );

      const compressedSecondFiles = await Promise.all(
        (student.secondFiles || []).map(async (f) => ({
          ...f,
          base64: await compressImageBase64(f.base64, f.type),
        }))
      );

      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          apiKey: settings.apiKey,
          model: settings.model,
          rubric1: {
            mode: settings.rubric1Mode,
            text: settings.rubric1Text,
            files: compressedRubric1Files,
          },
          rubric2: {
            mode: settings.rubric2Mode,
            text: settings.rubric2Text,
            files: compressedRubric2Files,
          },
          student: {
            name: student.name,
            firstFiles: compressedFirstFiles,
            secondFiles: compressedSecondFiles,
          }
        }),
      });

      let responseText = "";
      try {
        responseText = await response.text();
      } catch (err) {
        throw new Error("서버로부터 응답을 받지 못했습니다. 네트워크 연결을 확인해 주세요.");
      }

      let data: any;
      try {
        data = JSON.parse(responseText);
      } catch (err) {
        // Handle non-JSON responses (like Vercel Payload Too Large, 413 or general HTML gateway errors) gracefully
        if (response.status === 413 || responseText.includes("Too Large") || responseText.includes("Request Entity Too Large")) {
          throw new Error("전송된 이미지의 총 용량이 서버 제한(4.5MB)을 초과했습니다. 평가지 이미지 파일 크기를 압축하거나 장수를 줄여서 다시 시도해 주세요.");
        }
        console.error("Non-JSON Server Response:", responseText);
        throw new Error(`분석 중 오류가 발생했습니다. (서버 응답코드: ${response.status}). 관리자에게 문의해 주세요.`);
      }

      if (!response.ok || !data.success) {
        throw new Error(data.error || "분석 요청에 실패했습니다.");
      }

      // Save Successful result
      const analyzedStudent: StudentData = {
        ...student,
        analysis: {
          status: "done",
          result: data.result,
          analyzedAt: Date.now(),
        }
      };

      await dbService.saveStudent(analyzedStudent);
      setStudents(prev => prev.map(s => s.id === studentId ? analyzedStudent : s));
      
      if (!quiet) {
        showToast(`${student.name} 학생의 성장이 성공적으로 분석되었습니다!`, "success");
      }
      return true;

    } catch (err: any) {
      console.error(err);
      const failedStudent: StudentData = {
        ...student,
        analysis: {
          status: "error",
          error: err.message || "알 수 없는 API 에러가 발생했습니다.",
          analyzedAt: Date.now(),
        }
      };
      await dbService.saveStudent(failedStudent);
      setStudents(prev => prev.map(s => s.id === studentId ? failedStudent : s));
      
      if (!quiet) {
        showToast(`${student.name} 학생 분석 실패: ${err.message}`, "error");
      }
      return false;
    } finally {
      setApiProcessing((prev) => ({ ...prev, [studentId]: false }));
    }
  };

  const analyzeAllStudentsSequentially = async () => {
    // Collect students eligible for analysis (have both 1st and 2nd round files)
    const analyzable = students.filter(
      (s) => s.firstFiles.length > 0 && s.secondFiles.length > 0 && s.analysis?.status !== "running"
    );

    if (analyzable.length === 0) {
      showToast("분석 가능한 평가지(1차 및 2차 모두 업로드 완료)를 가진 학생이 없습니다.", "error");
      return;
    }

    const hasRubric1 = settings.rubric1Mode === "text" ? settings.rubric1Text.trim() : settings.rubric1Files.length > 0;
    const hasRubric2 = settings.rubric2Mode === "text" ? settings.rubric2Text.trim() : settings.rubric2Files.length > 0;

    if (!hasRubric1 || !hasRubric2) {
      showToast("설정 탭에서 1차 및 2차 평가 루브릭을 먼저 작성하거나 업로드해 주세요.", "error");
      return;
    }

    setIsAnalyzingAll(true);
    showToast(`${analyzable.length}명의 대기 학생에 대한 일괄 분석을 순차적으로 시작합니다. (Rate limit 방지를 위해 2초 간격 작동)`, "info");

    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < analyzable.length; i++) {
      const student = analyzable[i];
      const success = await analyzeSingleStudent(student.id, true);
      if (success) {
        successCount++;
      } else {
        failCount++;
      }

      // Pause to avoid API rate limit
      if (i < analyzable.length - 1) {
        await new Promise((resolve) => setTimeout(resolve, 2000));
      }
    }

    setIsAnalyzingAll(false);
    showToast(`일괄 분석 완료! 성공: ${successCount}명, 실패: ${failCount}명`, successCount > 0 ? "success" : "error");
  };

  // --- Global Clears & Export ---
  const handleClearAllData = async () => {
    try {
      await dbService.clearAllData();
      setSettings({
        id: "main",
        apiKey: "",
        model: "gemini-2.5-flash",
        rememberKey: false,
        rubric1Mode: "file",
        rubric1Text: "",
        rubric1Files: [],
        rubric2Mode: "file",
        rubric2Text: "",
        rubric2Files: [],
      });
      setStudents([]);
      setSelectedStudentId(null);
      setShowClearConfirmModal(false);
      showToast("데이터베이스의 모든 설정과 학생 기록이 초기화되었습니다.", "success");
    } catch (err) {
      showToast("초기화 처리 중 에러가 발생했습니다.", "error");
    }
  };

  const handleExportJson = () => {
    try {
      const exportData = {
        settings,
        students,
        exportedAt: new Date().toISOString(),
      };
      const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `student_growth_dashboard_data_${new Date().toISOString().split("T")[0]}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast("분석 결과 백업 데이터가 성공적으로 다운로드되었습니다.", "success");
    } catch (err) {
      showToast("데이터 내보내기에 실패했습니다.", "error");
    }
  };

  // --- Calculations for Analytics ---
  const completedStudents = students.filter((s) => s.analysis?.status === "done" && s.analysis.result);
  
  const calculateClassStats = () => {
    if (completedStudents.length === 0) {
      return {
        count: 0,
        k: { first: 0, second: 0, diff: 0 },
        p: { first: 0, second: 0, diff: 0 },
        v: { first: 0, second: 0, diff: 0 },
      };
    }

    let kFirstSum = 0, kFirstCount = 0;
    let kSecondSum = 0, kSecondCount = 0;
    let pFirstSum = 0, pFirstCount = 0;
    let pSecondSum = 0, pSecondCount = 0;
    let vFirstSum = 0, vFirstCount = 0;
    let vSecondSum = 0, vSecondCount = 0;

    completedStudents.forEach((std) => {
      const res = std.analysis!.result as AnalysisResult;
      
      // 지식이해
      if (res.first_scores.지식이해.normalized !== null) {
        kFirstSum += res.first_scores.지식이해.normalized;
        kFirstCount++;
      }
      if (res.second_scores.지식이해.normalized !== null) {
        kSecondSum += res.second_scores.지식이해.normalized;
        kSecondCount++;
      }

      // 과정기능
      if (res.first_scores.과정기능.normalized !== null) {
        pFirstSum += res.first_scores.과정기능.normalized;
        pFirstCount++;
      }
      if (res.second_scores.과정기능.normalized !== null) {
        pSecondSum += res.second_scores.과정기능.normalized;
        pSecondCount++;
      }

      // 가치태도
      if (res.first_scores.가치태도.normalized !== null) {
        vFirstSum += res.first_scores.가치태도.normalized;
        vFirstCount++;
      }
      if (res.second_scores.가치태도.normalized !== null) {
        vSecondSum += res.second_scores.가치태도.normalized;
        vSecondCount++;
      }
    });

    const kFirst = kFirstCount > 0 ? kFirstSum / kFirstCount : 0;
    const kSecond = kSecondCount > 0 ? kSecondSum / kSecondCount : 0;

    const pFirst = pFirstCount > 0 ? pFirstSum / pFirstCount : 0;
    const pSecond = pSecondCount > 0 ? pSecondSum / pSecondCount : 0;

    const vFirst = vFirstCount > 0 ? vFirstSum / vFirstCount : 0;
    const vSecond = vSecondCount > 0 ? vSecondSum / vSecondCount : 0;

    return {
      count: completedStudents.length,
      k: { first: kFirst, second: kSecond, diff: kSecond - kFirst },
      p: { first: pFirst, second: pSecond, diff: pSecond - pFirst },
      v: { first: vFirst, second: vSecond, diff: vSecond - vFirst },
    };
  };

  const classStats = calculateClassStats();

  const selectedStudent = students.find((s) => s.id === selectedStudentId);
  const selectedResult = selectedStudent?.analysis?.status === "done" ? (selectedStudent.analysis.result as AnalysisResult) : null;

  // --- Radar Chart SVG Generator ---
  const renderRadarChart = (result: AnalysisResult) => {
    // 3 axes angles: 0 is Up (-90deg), 1 is Right-Bottom (30deg), 2 is Left-Bottom (150deg)
    const center = 110;
    const maxRadius = 80;
    const valueMultiplier = maxRadius / 5; // 16px per score point (0 to 5)

    // Helper to get coordinates
    const getCoords = (val: number | null, index: number) => {
      const score = val === null ? 0 : val;
      const r = score * valueMultiplier;
      let angle = 0;
      if (index === 0) angle = -Math.PI / 2; // Up
      else if (index === 1) angle = Math.PI / 6; // Right-Bottom (30 deg)
      else if (index === 2) angle = (5 * Math.PI) / 6; // Left-Bottom (150 deg)

      return {
        x: center + r * Math.cos(angle),
        y: center + r * Math.sin(angle),
      };
    };

    // Score extraction
    const firstK = result.first_scores.지식이해.normalized;
    const firstP = result.first_scores.과정기능.normalized;
    const firstV = result.first_scores.가치태도.normalized;

    const secondK = result.second_scores.지식이해.normalized;
    const secondP = result.second_scores.과정기능.normalized;
    const secondV = result.second_scores.가치태도.normalized;

    // Get polygon points
    const p1_1 = getCoords(firstK, 0);
    const p1_2 = getCoords(firstP, 1);
    const p1_3 = getCoords(firstV, 2);

    const p2_1 = getCoords(secondK, 0);
    const p2_2 = getCoords(secondP, 1);
    const p2_3 = getCoords(secondV, 2);

    const firstPoints = `${p1_1.x},${p1_1.y} ${p1_2.x},${p1_2.y} ${p1_3.x},${p1_3.y}`;
    const secondPoints = `${p2_1.x},${p2_1.y} ${p2_2.x},${p2_2.y} ${p2_3.x},${p2_3.y}`;

    return (
      <svg viewBox="0 0 220 220" className="w-full h-full">
        {/* Web Concentric Guidelines (0 to 5) */}
        {[1, 2, 3, 4, 5].map((level) => {
          const r = level * valueMultiplier;
          // Draw triangles instead of circles to match 3 axes web grid
          const pt1 = getCoords(level, 0);
          const pt2 = getCoords(level, 1);
          const pt3 = getCoords(level, 2);
          return (
            <polygon
              key={level}
              points={`${pt1.x},${pt1.y} ${pt2.x},${pt2.y} ${pt3.x},${pt3.y}`}
              fill="none"
              stroke="#e8e5dc"
              strokeDasharray="3"
              strokeWidth="1"
            />
          );
        })}

        {/* 3 Axes Lines */}
        {[0, 1, 2].map((idx) => {
          const end = getCoords(5, idx);
          return (
            <line
              key={idx}
              x1={center}
              y1={center}
              x2={end.x}
              y2={end.y}
              stroke="#d4d0c2"
              strokeWidth="1.5"
            />
          );
        })}

        {/* 1st Evaluation Polygon (Pale Leaf / Warm) */}
        <polygon
          points={firstPoints}
          fill="var(--color-leaf-200)"
          fillOpacity="0.45"
          stroke="var(--color-leaf-400)"
          strokeWidth="1.5"
          className="transition-all duration-300"
        />

        {/* 2nd Evaluation Polygon (Dark Forest / Growth) */}
        <polygon
          points={secondPoints}
          fill="var(--color-leaf-600)"
          fillOpacity="0.25"
          stroke="var(--color-leaf-600)"
          strokeWidth="2.5"
          className="transition-all duration-300"
        />

        {/* Data Points - 1st Round */}
        <circle cx={p1_1.x} cy={p1_1.y} r="3.5" fill="var(--color-leaf-400)" />
        <circle cx={p1_2.x} cy={p1_2.y} r="3.5" fill="var(--color-leaf-400)" />
        <circle cx={p1_3.x} cy={p1_3.y} r="3.5" fill="var(--color-leaf-400)" />

        {/* Data Points - 2nd Round */}
        <circle cx={p2_1.x} cy={p2_1.y} r="4.5" fill="var(--color-leaf-600)" stroke="#ffffff" strokeWidth="1" />
        <circle cx={p2_2.x} cy={p2_2.y} r="4.5" fill="var(--color-leaf-600)" stroke="#ffffff" strokeWidth="1" />
        <circle cx={p2_3.x} cy={p2_3.y} r="4.5" fill="var(--color-leaf-600)" stroke="#ffffff" strokeWidth="1" />

        {/* Score Numbers Label */}
        {[1, 2, 3, 4, 5].map((level) => {
          const pt = getCoords(level, 0);
          return (
            <text
              key={level}
              x={pt.x + 6}
              y={pt.y + 4}
              fontSize="8"
              className="fill-ink-muted font-sans"
            >
              {level}
            </text>
          );
        })}

        {/* Axis Titles (Placed slightly outer) */}
        <text x="110" y="16" textAnchor="middle" className="font-serif font-semibold fill-ink-soft text-xs">
          지식·이해 {firstK === null && <tspan className="fill-ink-muted text-[10px] font-normal"> (N/A)</tspan>}
        </text>
        <text x="184" y="168" textAnchor="middle" className="font-serif font-semibold fill-ink-soft text-xs">
          과정·기능 {firstP === null && <tspan className="fill-ink-muted text-[10px] font-normal"> (N/A)</tspan>}
        </text>
        <text x="36" y="168" textAnchor="middle" className="font-serif font-semibold fill-ink-soft text-xs">
          가치·태도 {firstV === null && <tspan className="fill-ink-muted text-[10px] font-normal"> (N/A)</tspan>}
        </text>
      </svg>
    );
  };

  return (
    <div className="min-h-screen flex flex-col bg-paper text-ink transition-colors duration-200">
      
      {/* Toast Banner Alerts */}
      <div className="fixed top-6 right-6 z-[9999] flex flex-col gap-2 max-w-sm pointer-events-none no-print">
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`pointer-events-auto p-4 rounded-lg shadow-md border flex items-start gap-3 transition-all transform translate-y-0 animate-fade-in ${
              t.type === "success"
                ? "bg-leaf-50 border-leaf-200 text-leaf-700"
                : t.type === "error"
                ? "bg-coral-soft border-coral text-coral"
                : "bg-amber-soft border-amber text-amber"
            }`}
          >
            {t.type === "success" && <CheckCircle2 className="w-5 h-5 shrink-0 mt-0.5" />}
            {t.type === "error" && <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />}
            {t.type === "info" && <FileText className="w-5 h-5 shrink-0 mt-0.5" />}
            <div className="text-sm font-medium leading-relaxed">{t.message}</div>
          </div>
        ))}
      </div>

      {/* --- Header (Sticky & Responsive) --- */}
      <header id="app-header" className="sticky top-0 z-50 bg-[#faf9f4]/95 backdrop-blur-md border-b border-line px-6 pt-5 pb-4 no-print transition-all duration-200">
        <div className="max-w-[1180px] mx-auto flex flex-col md:flex-row md:justify-between md:items-end gap-4">
          <div className="flex flex-col">
            <h1 className="font-serif-title font-bold text-2xl md:text-[26px] leading-tight flex items-baseline gap-2 pt-1">
              송운 AI활용 평가 분석 시스템
            </h1>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => window.print()}
              disabled={completedStudents.length === 0}
              className={`text-[13px] px-3.5 py-2 rounded-[6px] flex items-center gap-1.5 font-medium transition-opacity cursor-pointer ${
                completedStudents.length === 0 
                  ? "bg-ink-faint text-white opacity-50 cursor-not-allowed" 
                  : "bg-leaf-600 hover:bg-leaf-700 text-white"
              }`}
            >
              <Printer className="w-4 h-4" />
              <span>PDF로 인쇄</span>
            </button>
          </div>
        </div>
      </header>

      {/* --- Tab Navigation (Sticky & Sleek) --- */}
      <nav id="app-tabs" className="sticky top-[86px] z-40 bg-[#faf9f4]/95 backdrop-blur-md px-6 border-b border-line no-print transition-all duration-200">
        <div className="max-w-[1180px] mx-auto flex gap-8">
          <button
            onClick={() => setActiveTab("settings")}
            className={`py-3.5 border-b-2 flex items-center gap-2 font-semibold text-sm transition-colors cursor-pointer ${
              activeTab === "settings"
                ? "border-leaf-600 text-leaf-600"
                : "border-transparent text-ink-muted hover:text-ink-soft"
            }`}
          >
            <span className="font-serif italic text-xs text-ink-muted">01</span>
            <Settings className="w-4 h-4" />
            <span>설정</span>
          </button>
          
          <button
            onClick={() => setActiveTab("students")}
            className={`py-3.5 border-b-2 flex items-center gap-2 font-semibold text-sm transition-colors cursor-pointer relative ${
              activeTab === "students"
                ? "border-leaf-600 text-leaf-600"
                : "border-transparent text-ink-muted hover:text-ink-soft"
            }`}
          >
            <span className="font-serif italic text-xs text-ink-muted">02</span>
            <Users className="w-4 h-4" />
            <span>학생 평가지</span>
            {students.length > 0 && (
              <span className="ml-1 px-1.5 py-0.2 bg-leaf-100 text-leaf-700 text-[10px] rounded-full font-bold">
                {students.length}
              </span>
            )}
          </button>

          <button
            onClick={() => {
              setActiveTab("dashboard");
              // Auto-select a completed student if nothing is selected
              if (!selectedStudentId && completedStudents.length > 0) {
                setSelectedStudentId(completedStudents[0].id);
              } else if (!selectedStudentId && students.length > 0) {
                setSelectedStudentId(students[0].id);
              }
            }}
            className={`py-3.5 border-b-2 flex items-center gap-2 font-semibold text-sm transition-colors cursor-pointer ${
              activeTab === "dashboard"
                ? "border-leaf-600 text-leaf-600"
                : "border-transparent text-ink-muted hover:text-ink-soft"
            }`}
          >
            <span className="font-serif italic text-xs text-ink-muted">03</span>
            <BarChart3 className="w-4 h-4" />
            <span>성장 대시보드</span>
            {completedStudents.length > 0 && (
              <span className="ml-1 px-1.5 py-0.2 bg-amber-soft text-amber border border-amber/30 text-[10px] rounded-full font-bold">
                {completedStudents.length} 완료
              </span>
            )}
          </button>
        </div>
      </nav>

      {/* --- Main Content Section --- */}
      <main className="flex-1 w-full max-w-[1180px] mx-auto px-6 py-6 transition-all duration-200">
        
        {/* ==================== 01 설정 탭 ==================== */}
        {activeTab === "settings" && (
          <div className="flex flex-col gap-6 animate-fade-in no-print">
            
            {/* 상단 안내 박스 */}
            <div className="p-5 bg-leaf-50 border-l-4 border-leaf-400 rounded-r-lg flex items-start gap-4">
              <FileText className="w-6 h-6 text-leaf-600 shrink-0 mt-0.5" />
              <div className="flex flex-col gap-1.5 text-sm">
                <h4 className="font-serif font-bold text-leaf-700 text-base">교사용 분석 솔루션 가이드</h4>
                <p className="text-ink-soft leading-relaxed">
                  이 도구는 교사가 1차 및 2차 서·논술형 평가 루브릭과 학생들의 필기 평가지를 업로드하면, AI 분석 기능을 활용해 지식·이해, 과정·기능, 가치·태도 영역의 성장을 분석합니다. 모든 데이터는 브라우저 내부 데이터베이스에 보안 저장을 원칙으로 하며, 외부에 전송되지 않습니다. AI분석에 대한 교사의 검토를 필수적으로 진행 합니다.
                </p>
              </div>
            </div>

            {/* 카드 2: 1차 평가 루브릭 */}
            <div className="bg-white border border-line rounded-lg p-6 shadow-sm">
              <div className="flex flex-col md:flex-row md:justify-between md:items-center gap-4 mb-4">
                <div>
                  <h3 className="font-serif text-lg font-semibold text-ink flex items-center gap-2">
                    <span className="serif italic text-xs bg-leaf-50 text-leaf-600 px-1.5 py-0.5 rounded border border-leaf-200">
                      1st ROUND
                    </span>
                    1차 평가 루브릭
                  </h3>
                </div>
                
                {/* 작은 약 알약 형태의 탭 토글 */}
                <div className="flex bg-leaf-50 p-1 rounded-full border border-leaf-100 max-w-fit">
                  <button
                    onClick={() => handleRubricModeChange(1, "file")}
                    className={`px-3 py-1 rounded-full text-xs font-medium transition-all cursor-pointer ${
                      settings.rubric1Mode === "file"
                        ? "bg-leaf-600 text-white shadow-xs"
                        : "text-ink-muted hover:text-ink-soft"
                    }`}
                  >
                    문서/이미지 업로드
                  </button>
                  <button
                    onClick={() => handleRubricModeChange(1, "text")}
                    className={`px-3 py-1 rounded-full text-xs font-medium transition-all cursor-pointer ${
                      settings.rubric1Mode === "text"
                        ? "bg-leaf-600 text-white shadow-xs"
                        : "text-ink-muted hover:text-ink-soft"
                    }`}
                  >
                    텍스트 직접 작성
                  </button>
                </div>
              </div>

              {settings.rubric1Mode === "text" ? (
                <textarea
                  value={settings.rubric1Text}
                  onChange={(e) => handleRubricTextChange(1, e.target.value)}
                  placeholder="예: [4수04-01] 막대그래프 서논술형 채점기준 표
1. 가로, 세로축 라벨과 제목을 모두 적었는가? (지식·이해, 만점 2점)
2. 표의 성격을 보고 적당한 눈금 한 칸 단위를 지정하여 그래프를 올바르게 그렸는가? (과정·기능, 만점 3점)
3. 그래프의 특징을 발견하고, 생활속에서 활용할 수 있는 아이디어를 제시하였는가? (가치·태도, 만점 2점)"
                  className="w-full min-h-[140px] px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 focus:border-leaf-600 font-mono bg-[#fafafa]"
                />
              ) : (
                <div className="flex flex-col gap-3">
                  <div className="border-2 border-dashed border-line-strong hover:border-leaf-400 rounded-lg p-6 bg-leaf-50/20 text-center transition-colors cursor-pointer relative">
                    <input
                      type="file"
                      multiple
                      accept="application/pdf,image/*"
                      onChange={(e) => handleRubricFileUpload(1, e)}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    />
                    <Upload className="w-8 h-8 text-leaf-600 mx-auto mb-2" />
                    <span className="text-xs font-semibold text-ink-soft block">
                      이곳을 클릭하거나 문서를 끌어다 놓으세요.
                    </span>
                    <span className="text-[10px] text-ink-muted mt-1 block">
                      PDF, JPG, PNG 파일 형식 지원 (다중 선택 가능)
                    </span>
                  </div>

                  {settings.rubric1Files.length > 0 && (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2 mt-2">
                      {settings.rubric1Files.map((file) => (
                        <div key={file.id} className="flex justify-between items-center p-2.5 bg-white border border-line rounded-md text-xs">
                          <div className="flex items-center gap-2 overflow-hidden mr-2">
                            <span className="text-leaf-600 font-bold shrink-0">📄</span>
                            <span className="truncate text-ink-soft font-medium" title={file.name}>{file.name}</span>
                            <span className="text-[10px] text-ink-muted shrink-0">
                              ({(file.size / 1024).toFixed(0)}KB)
                            </span>
                          </div>
                          <button
                            onClick={() => handleDeleteRubricFile(1, file.id)}
                            className="text-coral hover:bg-coral-soft p-1 rounded transition-colors cursor-pointer shrink-0"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* 카드 3: 2차 평가 루브릭 */}
            <div className="bg-white border border-line rounded-lg p-6 shadow-sm">
              <div className="flex flex-col md:flex-row md:justify-between md:items-center gap-4 mb-4">
                <div>
                  <h3 className="font-serif text-lg font-semibold text-ink flex items-center gap-2">
                    <span className="serif italic text-xs bg-amber-soft text-amber px-1.5 py-0.5 rounded border border-amber/30">
                      2nd ROUND
                    </span>
                    2차 평가 루브릭
                  </h3>
                </div>
                
                <div className="flex bg-leaf-50 p-1 rounded-full border border-leaf-100 max-w-fit">
                  <button
                    onClick={() => handleRubricModeChange(2, "file")}
                    className={`px-3 py-1 rounded-full text-xs font-medium transition-all cursor-pointer ${
                      settings.rubric2Mode === "file"
                        ? "bg-leaf-600 text-white shadow-xs"
                        : "text-ink-muted hover:text-ink-soft"
                    }`}
                  >
                    문서/이미지 업로드
                  </button>
                  <button
                    onClick={() => handleRubricModeChange(2, "text")}
                    className={`px-3 py-1 rounded-full text-xs font-medium transition-all cursor-pointer ${
                      settings.rubric2Mode === "text"
                        ? "bg-leaf-600 text-white shadow-xs"
                        : "text-ink-muted hover:text-ink-soft"
                    }`}
                  >
                    텍스트 직접 작성
                  </button>
                </div>
              </div>

              {settings.rubric2Mode === "text" ? (
                <textarea
                  value={settings.rubric2Text}
                  onChange={(e) => handleRubricTextChange(2, e.target.value)}
                  placeholder="예: [4수04-01] 막대그래프 심화 서논술형 채점기준 표
1. 가로축, 세로축 명칭 및 그래프 제목의 누락이 없는가? (지식·이해, 만점 2점)
2. 대용량 수치 변동에 적절한 눈금 단위(예: 한 눈금 2 또는 5)를 설정하고, 막대의 길이를 정밀하게 표현했는가? (과정·기능, 만점 5점)
3. 수집된 결과에 대해 의사결정적 해석을 내리고 유의미한 가치적 의견을 달았는가? (가치·태도, 만점 3점)"
                  className="w-full min-h-[140px] px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 focus:border-leaf-600 font-mono bg-[#fafafa]"
                />
              ) : (
                <div className="flex flex-col gap-3">
                  <div className="border-2 border-dashed border-line-strong hover:border-leaf-400 rounded-lg p-6 bg-leaf-50/20 text-center transition-colors cursor-pointer relative">
                    <input
                      type="file"
                      multiple
                      accept="application/pdf,image/*"
                      onChange={(e) => handleRubricFileUpload(2, e)}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    />
                    <Upload className="w-8 h-8 text-leaf-600 mx-auto mb-2" />
                    <span className="text-xs font-semibold text-ink-soft block">
                      이곳을 클릭하거나 문서를 끌어다 놓으세요.
                    </span>
                    <span className="text-[10px] text-ink-muted mt-1 block">
                      PDF, JPG, PNG 파일 형식 지원 (다중 선택 가능)
                    </span>
                  </div>

                  {settings.rubric2Files.length > 0 && (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2 mt-2">
                      {settings.rubric2Files.map((file) => (
                        <div key={file.id} className="flex justify-between items-center p-2.5 bg-white border border-line rounded-md text-xs">
                          <div className="flex items-center gap-2 overflow-hidden mr-2">
                            <span className="text-leaf-600 font-bold shrink-0">📄</span>
                            <span className="truncate text-ink-soft font-medium" title={file.name}>{file.name}</span>
                            <span className="text-[10px] text-ink-muted shrink-0">
                              ({(file.size / 1024).toFixed(0)}KB)
                            </span>
                          </div>
                          <button
                            onClick={() => handleDeleteRubricFile(2, file.id)}
                            className="text-coral hover:bg-coral-soft p-1 rounded transition-colors cursor-pointer shrink-0"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* 카드 4: 데이터 관리 */}
            <div className="bg-white border border-line rounded-lg p-6 shadow-sm border-l-4 border-coral">
              <h3 className="font-serif text-lg font-semibold mb-1 text-ink flex items-center gap-2">
                학급 데이터 관리 및 보존
              </h3>
              <p className="text-xs text-ink-muted mb-4">
                이 브라우저 로컬 데이터베이스를 제어하며 전체 학생 평가지와 루브릭 세팅을 완전 제거하거나 백업할 수 있습니다.
              </p>
              
              <div className="flex flex-wrap gap-3">
                <button
                  onClick={handleExportJson}
                  className="text-xs border border-line-strong bg-white hover:bg-leaf-50 px-4 py-2.5 rounded font-medium flex items-center gap-2 cursor-pointer"
                >
                  <Download className="w-4 h-4 text-leaf-600" />
                  <span>분석 결과 JSON 다운로드 (백업)</span>
                </button>
                <button
                  onClick={() => setShowClearConfirmModal(true)}
                  className="text-xs border border-coral text-coral bg-white hover:bg-coral-soft px-4 py-2.5 rounded font-medium flex items-center gap-2 cursor-pointer"
                >
                  <Trash2 className="w-4 h-4" />
                  <span>학급 데이터 전체 초기화 (IndexedDB 포맷)</span>
                </button>
              </div>
            </div>

          </div>
        )}

        {/* ==================== 02 학생 평가지 탭 ==================== */}
        {activeTab === "students" && (
          <div className="flex flex-col gap-5 animate-fade-in no-print">
            
            {/* 툴바 */}
            <div className="bg-white border border-line rounded-lg p-4 flex flex-col sm:flex-row sm:justify-between sm:items-center gap-4 shadow-xs">
              <div className="flex items-center gap-3">
                <div className="text-sm font-semibold text-ink-soft">
                  학급 등록 상태: <span className="font-serif text-base text-leaf-600">{students.length}명</span>
                </div>
                <div className="h-4 w-[1px] bg-line" />
                <div className="flex gap-2 text-xs">
                  <span className="bg-leaf-50 text-leaf-700 px-2.5 py-0.5 rounded-full font-medium">
                    분석 완료: {students.filter((s) => s.analysis?.status === "done").length}명
                  </span>
                  <span className="bg-amber-soft text-amber px-2.5 py-0.5 rounded-full font-medium">
                    대기/진행: {students.filter((s) => s.analysis?.status === "running" || (!s.analysis && s.firstFiles.length > 0 && s.secondFiles.length > 0)).length}명
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={analyzeAllStudentsSequentially}
                  disabled={isAnalyzingAll || students.length === 0}
                  className={`text-xs px-3.5 py-2.5 rounded font-semibold flex items-center gap-1.5 cursor-pointer transition-colors ${
                    isAnalyzingAll 
                      ? "bg-amber-soft text-amber cursor-not-allowed border border-amber/30" 
                      : "bg-leaf-600 hover:bg-leaf-700 text-white"
                  }`}
                >
                  {isAnalyzingAll ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>순차 분석 처리 중...</span>
                    </>
                  ) : (
                    <>
                      <Play className="w-4 h-4" />
                      <span>대기 학생 모두 분석 실행</span>
                    </>
                  )}
                </button>

                <button
                  onClick={() => setShowAddModal(true)}
                  className="text-xs border border-line-strong bg-white hover:bg-leaf-50 px-3.5 py-2.5 rounded font-semibold flex items-center gap-1.5 cursor-pointer transition-colors"
                >
                  <Plus className="w-4 h-4 text-leaf-600" />
                  <span>학생 추가 등록</span>
                </button>
              </div>
            </div>

            {/* 학생 목록 그리드 */}
            {students.length === 0 ? (
              <div className="bg-white border border-line rounded-lg p-12 text-center text-ink-soft">
                <Users className="w-12 h-12 text-ink-muted mx-auto mb-3" />
                <h4 className="font-serif font-bold text-lg mb-1 text-ink">아직 등록된 학생이 없습니다.</h4>
                <p className="text-xs text-ink-muted mb-4">
                  평가 및 분석을 시작하기 위해 첫 번째 학생을 학급 명단에 등록하세요.
                </p>
                <button
                  onClick={() => setShowAddModal(true)}
                  className="text-xs bg-leaf-600 text-white px-4 py-2.5 rounded font-semibold inline-flex items-center gap-1.5 hover:opacity-90 transition-opacity cursor-pointer"
                >
                  <Plus className="w-4 h-4" />
                  <span>새로운 학생 등록</span>
                </button>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                {students.map((student) => {
                  // status computation
                  let statusText = "파일 필요";
                  let statusStyle = "bg-ink-faint text-ink-soft";
                  const hasFiles = student.firstFiles.length > 0 && student.secondFiles.length > 0;

                  if (student.analysis?.status === "running") {
                    statusText = "분석 중";
                    statusStyle = "bg-amber-soft text-amber animate-pulse border border-amber/30";
                  } else if (student.analysis?.status === "done") {
                    statusText = "분석 완료";
                    statusStyle = "bg-leaf-100 text-leaf-700 border border-leaf-200";
                  } else if (student.analysis?.status === "error") {
                    statusText = "분석 오류";
                    statusStyle = "bg-coral-soft text-coral border border-coral/30";
                  } else if (hasFiles) {
                    statusText = "분석 가능";
                    statusStyle = "bg-leaf-50 text-leaf-600 border border-leaf-200";
                  }

                  const isProcessing = apiProcessing[student.id];

                  return (
                    <div
                      key={student.id}
                      className="bg-white border border-line rounded-lg p-5 flex flex-col gap-4 shadow-xs relative hover:border-leaf-200 transition-colors"
                    >
                      {/* 카드 상단 헤더 */}
                      <div className="flex justify-between items-start border-b border-line pb-3">
                        <div className="flex items-center gap-3">
                          <div className="flex flex-col">
                            <span className="font-serif text-lg font-bold text-ink">
                              {student.name}
                            </span>
                            <span className="text-[11px] text-ink-muted font-medium">
                              {student.className} · {student.number ? `${student.number}번` : "번호 미지정"}
                            </span>
                          </div>
                          <span className={`text-[11px] px-2.5 py-0.5 rounded-full font-semibold ${statusStyle}`}>
                            {statusText}
                          </span>
                        </div>

                        <div className="flex items-center gap-1.5">
                          <button
                            onClick={() => analyzeSingleStudent(student.id)}
                            disabled={isProcessing || isAnalyzingAll || !hasFiles}
                            title={hasFiles ? "분석하기" : "1/2차 평가지를 먼저 올려야 분석이 가능합니다."}
                            className={`p-1.5 rounded transition-colors cursor-pointer shrink-0 ${
                              hasFiles && !isProcessing && !isAnalyzingAll
                                ? "text-leaf-600 hover:bg-leaf-50"
                                : "text-ink-muted opacity-40 cursor-not-allowed"
                            }`}
                          >
                            {isProcessing ? (
                              <Loader2 className="w-4 h-4 animate-spin" />
                            ) : (
                              <Play className="w-4 h-4" />
                            )}
                          </button>
                          
                          <button
                            onClick={() => handleEditStudentClick(student)}
                            className="text-ink-soft hover:bg-leaf-50 p-1.5 rounded transition-colors cursor-pointer shrink-0"
                            title="정보 수정"
                          >
                            <FileText className="w-4 h-4" />
                          </button>
                          
                          <button
                            onClick={() => handleDeleteStudent(student.id, student.name)}
                            className="text-coral hover:bg-coral-soft p-1.5 rounded transition-colors cursor-pointer shrink-0"
                            title="완전 삭제"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </div>

                      {/* 1차 vs 2차 평가지 업로드 더블 영역 */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        {/* 1차 평가지 업로드 영역 */}
                        <div className="flex flex-col gap-2">
                          <span className="text-xs font-semibold text-leaf-600">
                            1차 평가지
                          </span>
                          <div
                            onDragOver={(e) => handleDragOver(e, `${student.id}-1`)}
                            onDragLeave={handleDragLeave}
                            onDrop={(e) => handleDrop(e, student.id, 1)}
                            className={`border border-dashed rounded-md p-3 text-center transition-all relative ${
                              dragActiveId === `${student.id}-1`
                                ? "border-leaf-600 bg-leaf-50/50"
                                : "border-line-strong hover:bg-leaf-50/20"
                            }`}
                          >
                            <input
                              type="file"
                              multiple
                              accept="image/*,application/pdf"
                              onChange={(e) => handleStudentFileUpload(student.id, 1, e)}
                              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                            />
                            <Upload className="w-5 h-5 text-leaf-400 mx-auto mb-1" />
                            <span className="text-[10px] font-semibold text-ink-soft block">
                              클릭 또는 이미지 추가
                            </span>
                          </div>

                          {student.firstFiles.length > 0 && (
                            <div className="flex flex-col gap-1 max-h-24 overflow-y-auto">
                              {student.firstFiles.map((f) => (
                                <div key={f.id} className="flex justify-between items-center bg-leaf-50/40 p-1 px-2 rounded border border-line text-[10px]">
                                  <span className="truncate text-ink-soft font-medium max-w-[110px]" title={f.name}>
                                    {f.name}
                                  </span>
                                  <button
                                    onClick={() => handleDeleteStudentFile(student.id, 1, f.id)}
                                    className="text-coral hover:text-coral font-bold ml-1 cursor-pointer text-xs"
                                  >
                                    ✕
                                  </button>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>

                        {/* 2차 평가지 업로드 영역 */}
                        <div className="flex flex-col gap-2">
                          <span className="text-xs font-semibold text-amber">
                            2차 평가지
                          </span>
                          <div
                            onDragOver={(e) => handleDragOver(e, `${student.id}-2`)}
                            onDragLeave={handleDragLeave}
                            onDrop={(e) => handleDrop(e, student.id, 2)}
                            className={`border border-dashed rounded-md p-3 text-center transition-all relative ${
                              dragActiveId === `${student.id}-2`
                                ? "border-amber bg-amber-soft/20"
                                : "border-line-strong hover:bg-leaf-50/20"
                            }`}
                          >
                            <input
                              type="file"
                              multiple
                              accept="image/*,application/pdf"
                              onChange={(e) => handleStudentFileUpload(student.id, 2, e)}
                              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                            />
                            <Upload className="w-5 h-5 text-amber mx-auto mb-1 opacity-75" />
                            <span className="text-[10px] font-semibold text-ink-soft block">
                              클릭 또는 이미지 추가
                            </span>
                          </div>

                          {student.secondFiles.length > 0 && (
                            <div className="flex flex-col gap-1 max-h-24 overflow-y-auto">
                              {student.secondFiles.map((f) => (
                                <div key={f.id} className="flex justify-between items-center bg-amber-soft/20 p-1 px-2 rounded border border-line text-[10px]">
                                  <span className="truncate text-ink-soft font-medium max-w-[110px]" title={f.name}>
                                    {f.name}
                                  </span>
                                  <button
                                    onClick={() => handleDeleteStudentFile(student.id, 2, f.id)}
                                    className="text-coral hover:text-coral font-bold ml-1 cursor-pointer text-xs"
                                  >
                                    ✕
                                  </button>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>

                      {/* 개별 분석 에러 로그 피드백 */}
                      {student.analysis?.status === "error" && (
                        <div className="mt-1 p-2 bg-coral-soft rounded border border-coral/20 text-[11px] text-coral flex items-start gap-1.5 leading-relaxed font-mono">
                          <AlertCircle className="w-3.5 h-3.5 mt-0.5 shrink-0" />
                          <span>오류 내용: {student.analysis.error}</span>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* ==================== 03 성장 대시보드 탭 ==================== */}
        {activeTab === "dashboard" && (
          <div className="flex flex-col gap-6 animate-fade-in print-grid">
            
            {/* 학급 종합 배너 (상단) */}
            <div className="bg-gradient-to-r from-leaf-700 to-leaf-600 rounded-xl p-6 text-white flex flex-col md:flex-row md:items-center md:justify-between gap-6 shadow-md print-card">
              <div className="flex flex-col gap-1">
                <h3 className="font-serif text-lg md:text-xl font-semibold">학급 종합 성장 리포트</h3>
                <p className="text-xs opacity-85">
                  총 {completedStudents.length}명의 분석이 완료된 학생 평균 데이터입니다.
                </p>
              </div>

              {completedStudents.length === 0 ? (
                <span className="text-xs font-semibold bg-white/10 px-4 py-2.5 rounded-lg border border-white/20">
                  평가 분석 결과가 존재하지 않아 학급 종합 데이터가 비활성화되었습니다.
                </span>
              ) : (
                <div className="flex gap-8 md:gap-12 flex-wrap">
                  <div className="text-center md:text-left">
                    <div className="text-[10px] uppercase tracking-wider opacity-75 mb-1 font-semibold">지식·이해 평균</div>
                    <div className="text-xl md:text-2xl font-bold font-serif-title">
                      {classStats.k.second.toFixed(1)} <span className="text-xs font-sans font-normal opacity-80">/5점</span>
                      <span className="text-[#ffd28c] ml-1.5 text-sm md:text-base">
                        ▲ +{classStats.k.diff.toFixed(1)}
                      </span>
                    </div>
                  </div>
                  
                  <div className="text-center md:text-left">
                    <div className="text-[10px] uppercase tracking-wider opacity-75 mb-1 font-semibold">과정·기능 평균</div>
                    <div className="text-xl md:text-2xl font-bold font-serif-title">
                      {classStats.p.second.toFixed(1)} <span className="text-xs font-sans font-normal opacity-80">/5점</span>
                      <span className="text-[#ffd28c] ml-1.5 text-sm md:text-base">
                        ▲ +{classStats.p.diff.toFixed(1)}
                      </span>
                    </div>
                  </div>

                  <div className="text-center md:text-left">
                    <div className="text-[10px] uppercase tracking-wider opacity-75 mb-1 font-semibold">가치·태도 평균</div>
                    <div className="text-xl md:text-2xl font-bold font-serif-title">
                      {classStats.v.second.toFixed(1)} <span className="text-xs font-sans font-normal opacity-80">/5점</span>
                      <span className="text-[#ffd28c] ml-1.5 text-sm md:text-base">
                        ▲ +{classStats.v.diff.toFixed(1)}
                      </span>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* 메인 뷰: 좌측 학생 명단 사이드바 + 우측 학생 성장 상세 카드 */}
            <div className="flex flex-col lg:flex-row gap-6 print-grid">
              
              {/* 좌측 학생 명단 사이드바 */}
              <aside className="w-full lg:w-[260px] flex flex-col gap-3 shrink-0 no-print">
                <div className="small-caps text-ink-muted mb-1 text-xs font-semibold uppercase tracking-wider">
                  학급 분석 명단 ({students.length}명)
                </div>

                {students.length === 0 ? (
                  <div className="p-4 rounded-lg bg-white border border-line text-xs text-ink-muted text-center leading-relaxed">
                    등록된 학생이 존재하지 않습니다.
                  </div>
                ) : (
                  <div className="flex lg:flex-col gap-2 overflow-x-auto lg:overflow-y-auto max-h-[500px] pr-1 pb-2">
                    {students.map((std) => {
                      const isSelected = std.id === selectedStudentId;
                      const isAnalyzed = std.analysis?.status === "done";
                      
                      return (
                        <div
                          key={std.id}
                          onClick={() => setSelectedStudentId(std.id)}
                          className={`min-w-[180px] lg:w-full p-4 rounded-lg border transition-all cursor-pointer select-none ${
                            isSelected
                              ? "bg-leaf-600 border-leaf-600 text-white shadow-sm"
                              : "bg-white border-line text-ink hover:bg-leaf-50"
                          }`}
                        >
                          <div className="flex justify-between items-center">
                            <span className="font-serif text-[17px] font-bold">
                              {std.name}
                            </span>
                            <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-semibold ${
                              isSelected
                                ? "bg-white/20 text-white"
                                : isAnalyzed
                                ? "bg-leaf-50 text-leaf-700"
                                : "bg-ink-faint text-ink-soft"
                            }`}>
                              {isAnalyzed ? "분석 완료" : "미분석"}
                            </span>
                          </div>
                          <div className={`text-[11px] mt-1.5 font-medium ${
                            isSelected ? "text-white/80" : "text-ink-muted"
                          }`}>
                            {std.className} · {std.number ? `${std.number}번` : "번호 미지정"}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </aside>

              {/* 우측 상세 분석 리포트 카드 */}
              <div className="flex-1 bg-white border border-line rounded-lg p-6 md:p-8 flex flex-col gap-6 shadow-xs print-card">
                
                {!selectedStudent ? (
                  <div className="py-20 text-center text-ink-soft">
                    <FileText className="w-12 h-12 text-ink-muted mx-auto mb-3" />
                    <h4 className="font-serif font-bold text-lg text-ink">선택된 학생이 없습니다.</h4>
                    <p className="text-xs text-ink-muted">좌측 학급 명단에서 성장 내용을 볼 학생을 클릭해 주세요.</p>
                  </div>
                ) : selectedStudent.analysis?.status !== "done" ? (
                  <div className="py-20 text-center text-ink-soft">
                    <FileText className="w-12 h-12 text-ink-muted mx-auto mb-3" />
                    <h4 className="font-serif font-bold text-lg text-ink mb-1">
                      {selectedStudent.name} 학생의 분석 결과가 없습니다.
                    </h4>
                    <p className="text-xs text-ink-muted mb-4">
                      평가지 업로드 후 분석을 실행하거나 완료될 때까지 기다려 주세요.
                    </p>
                    <button
                      onClick={() => {
                        setActiveTab("students");
                      }}
                      className="text-xs border border-line-strong hover:bg-leaf-50 bg-white px-4 py-2 rounded font-semibold inline-flex items-center gap-1.5 cursor-pointer no-print"
                    >
                      <span>평가지 업로드/분석 탭으로 이동</span>
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                ) : selectedResult ? (
                  isEditingDashboard && editDashboardData ? (
                    <div className="flex flex-col gap-6 animate-fade-in text-left">
                      {/* 상단 타이틀 */}
                      <div className="flex flex-col md:flex-row md:justify-between md:items-start gap-4 border-b border-line pb-4">
                        <div className="flex flex-col gap-1">
                          <h2 className="font-serif-title text-xl font-bold text-ink">
                            {selectedStudent.name} 학생 분석 결과 검토 및 수정
                          </h2>
                          <p className="text-xs text-ink-muted">AI가 도출한 채점 점수와 정성 평가 및 지도 제안을 검토하고 수정할 수 있습니다.</p>
                        </div>
                        <div className="flex gap-2">
                          <button
                            onClick={() => {
                              setIsEditingDashboard(false);
                              setEditDashboardData(null);
                            }}
                            className="text-xs border border-line-strong hover:bg-leaf-50 bg-white px-3 py-2 rounded font-semibold cursor-pointer"
                          >
                            취소
                          </button>
                          <button
                            onClick={handleSaveDashboardData}
                            className="text-xs bg-leaf-600 hover:bg-leaf-700 text-white px-4 py-2 rounded font-semibold shadow-xs cursor-pointer"
                          >
                            저장하기
                          </button>
                        </div>
                      </div>

                      {/* Section 1: 영역별 정량 평가 수정 */}
                      <div className="bg-leaf-50/20 p-5 rounded-lg border border-line flex flex-col gap-4">
                        <h3 className="font-serif text-sm font-bold text-leaf-700 flex items-center gap-1.5">
                          <CheckCircle2 className="w-4 h-4 text-leaf-600" />
                          01. 영역별 정량 평가 수정
                        </h3>

                        {/* 영역별 상세 점수 */}
                        <div className="flex flex-col gap-4 mt-2">
                          <span className="text-[11px] font-bold text-ink-soft">영역별 점수 설정 (환산 점수와 기준 점수를 수정하면 5점 척도 환산점수가 자동으로 반영됩니다)</span>
                          
                          {(["지식이해", "과정기능", "가치태도"] as const).map((domain) => {
                            const domainLabel = domain === "지식이해" ? "지식·이해" : domain === "과정기능" ? "과정·기능" : "가치·태도";
                            const fScore = editDashboardData.first_scores?.[domain] || { score: 0, max: 5, normalized: 0, evidence: "" };
                            const sScore = editDashboardData.second_scores?.[domain] || { score: 0, max: 5, normalized: 0, evidence: "" };
                            
                            return (
                              <div key={domain} className="bg-white border border-line rounded-lg p-4 flex flex-col gap-3">
                                <span className="text-xs font-bold text-leaf-600">{domainLabel} 영역</span>
                                
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                  {/* 1차 평가 */}
                                  <div className="p-3 bg-leaf-50/10 rounded border border-line/30 flex flex-col gap-2.5">
                                    <span className="text-[10px] font-bold text-ink-muted uppercase">1차 평가 결과</span>
                                    <div className="grid grid-cols-2 gap-3">
                                      <div>
                                        <label className="block text-[9px] text-ink-muted mb-1">환산 점수</label>
                                        <input
                                          type="number"
                                          step="any"
                                          value={fScore.score === "" ? "" : fScore.score}
                                          onChange={(e) => handleScoreChange("first_scores", domain, "score", e.target.value)}
                                          className="text-xs w-full bg-white border border-line rounded p-1.5 font-mono text-center"
                                        />
                                      </div>
                                      <div>
                                        <label className="block text-[9px] text-ink-muted mb-1">기준 점수</label>
                                        <input
                                          type="number"
                                          step="any"
                                          value={fScore.max === "" ? "" : fScore.max}
                                          onChange={(e) => handleScoreChange("first_scores", domain, "max", e.target.value)}
                                          className="text-xs w-full bg-white border border-line rounded p-1.5 font-mono text-center"
                                        />
                                      </div>
                                    </div>
                                    <div className="flex flex-col gap-1">
                                      <label className="block text-[9px] text-ink-muted">채점/평가 근거 (피드백)</label>
                                      <textarea
                                        rows={2}
                                        value={fScore.evidence || ""}
                                        onChange={(e) => handleScoreChange("first_scores", domain, "evidence", e.target.value)}
                                        className="text-xs w-full bg-white border border-line rounded px-2.5 py-1.5 leading-relaxed focus:ring-1 focus:ring-leaf-500 focus:outline-hidden"
                                        placeholder="1차 채점 근거 및 피드백을 입력하세요..."
                                      />
                                    </div>
                                  </div>

                                  {/* 2차 평가 */}
                                  <div className="p-3 bg-leaf-50/10 rounded border border-line/30 flex flex-col gap-2.5">
                                    <span className="text-[10px] font-bold text-ink-muted uppercase">2차 평가 결과</span>
                                    <div className="grid grid-cols-2 gap-3">
                                      <div>
                                        <label className="block text-[9px] text-ink-muted mb-1">환산 점수</label>
                                        <input
                                          type="number"
                                          step="any"
                                          value={sScore.score === "" ? "" : sScore.score}
                                          onChange={(e) => handleScoreChange("second_scores", domain, "score", e.target.value)}
                                          className="text-xs w-full bg-white border border-line rounded p-1.5 font-mono text-center"
                                        />
                                      </div>
                                      <div>
                                        <label className="block text-[9px] text-ink-muted mb-1">기준 점수</label>
                                        <input
                                          type="number"
                                          step="any"
                                          value={sScore.max === "" ? "" : sScore.max}
                                          onChange={(e) => handleScoreChange("second_scores", domain, "max", e.target.value)}
                                          className="text-xs w-full bg-white border border-line rounded p-1.5 font-mono text-center"
                                        />
                                      </div>
                                    </div>
                                    <div className="flex flex-col gap-1">
                                      <label className="block text-[9px] text-ink-muted">채점/평가 근거 (피드백)</label>
                                      <textarea
                                        rows={2}
                                        value={sScore.evidence || ""}
                                        onChange={(e) => handleScoreChange("second_scores", domain, "evidence", e.target.value)}
                                        className="text-xs w-full bg-white border border-line rounded px-2.5 py-1.5 leading-relaxed focus:ring-1 focus:ring-leaf-500 focus:outline-hidden"
                                        placeholder="2차 채점 근거 및 피드백을 입력하세요..."
                                      />
                                    </div>
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>

                      {/* Section 2: 성장 종합 서평 */}
                      <div className="flex flex-col gap-2">
                        <label className="text-xs font-bold text-leaf-700 flex items-center gap-1.5">
                          <CheckCircle2 className="w-4 h-4 text-leaf-600" />
                          02. 성장 종합 서평 수정
                        </label>
                        <textarea
                          value={editDashboardData.overall_summary || ""}
                          onChange={(e) => setEditDashboardData(prev => prev ? { ...prev, overall_summary: e.target.value } : null)}
                          rows={4}
                          className="text-xs w-full bg-white border border-line rounded p-3 leading-relaxed font-medium focus:ring-1 focus:ring-leaf-500 focus:outline-hidden"
                          placeholder="학습 전반에 대한 성장 서평을 입력하세요"
                        />
                      </div>

                      {/* Section 3: 영역별 학습 성장 정성 분석 */}
                      <div className="flex flex-col gap-3">
                        <label className="text-xs font-bold text-leaf-700 flex items-center gap-1.5">
                          <CheckCircle2 className="w-4 h-4 text-leaf-600" />
                          03. 영역별 학습 성장 정성 분석 수정
                        </label>
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                          <div className="flex flex-col gap-1.5">
                            <span className="text-[11px] font-bold text-leaf-600">지식·이해 변화</span>
                            <textarea
                              value={editDashboardData.growth_analysis?.지식이해 || ""}
                              onChange={(e) => setEditDashboardData(prev => prev ? {
                                ...prev,
                                growth_analysis: {
                                  ...prev.growth_analysis,
                                  지식이해: e.target.value
                                }
                              } : null)}
                              rows={6}
                              className="text-xs w-full bg-[#faf9f4] border border-line rounded p-2.5 leading-relaxed"
                              placeholder="지식·이해 영역의 성장과 발전 사항을 입력하세요"
                            />
                          </div>

                          <div className="flex flex-col gap-1.5">
                            <span className="text-[11px] font-bold text-leaf-600">과정·기능 변화</span>
                            <textarea
                              value={editDashboardData.growth_analysis?.과정기능 || ""}
                              onChange={(e) => setEditDashboardData(prev => prev ? {
                                ...prev,
                                growth_analysis: {
                                  ...prev.growth_analysis,
                                  과정기능: e.target.value
                                }
                              } : null)}
                              rows={6}
                              className="text-xs w-full bg-[#faf9f4] border border-line rounded p-2.5 leading-relaxed"
                              placeholder="과정·기능 영역의 성장과 발전 사항을 입력하세요"
                            />
                          </div>

                          <div className="flex flex-col gap-1.5">
                            <span className="text-[11px] font-bold text-leaf-600">가치·태도 변화</span>
                            <textarea
                              value={editDashboardData.growth_analysis?.가치태도 || ""}
                              onChange={(e) => setEditDashboardData(prev => prev ? {
                                ...prev,
                                growth_analysis: {
                                  ...prev.growth_analysis,
                                  가치태도: e.target.value
                                }
                              } : null)}
                              rows={6}
                              className="text-xs w-full bg-[#faf9f4] border border-line rounded p-2.5 leading-relaxed"
                              placeholder="가치·태도 영역의 성장과 발전 사항을 입력하세요"
                            />
                          </div>
                        </div>
                      </div>

                      {/* Section 4: 맞춤형 지도 전략 */}
                      <div className="flex flex-col gap-2">
                        <label className="text-xs font-bold text-leaf-700 flex items-center gap-1.5">
                          <CheckCircle2 className="w-4 h-4 text-leaf-600" />
                          04. 개인별 맞춤 다음 단계 학습 지도 전략 수정
                        </label>
                        <textarea
                          value={editDashboardData.teaching_feedback || ""}
                          onChange={(e) => setEditDashboardData(prev => prev ? { ...prev, teaching_feedback: e.target.value } : null)}
                          rows={3}
                          className="text-xs w-full bg-white border border-line rounded p-3 leading-relaxed font-medium text-amber-900 border-amber-200 bg-amber-50/20"
                          placeholder="학생의 향후 성장을 촉진하기 위한 맞춤 학습 지도 전략을 입력하세요"
                        />
                      </div>

                      {/* 하단 버튼 */}
                      <div className="flex justify-end gap-2 border-t border-line pt-4 mt-2">
                        <button
                          onClick={() => {
                            setIsEditingDashboard(false);
                            setEditDashboardData(null);
                          }}
                          className="text-xs border border-line-strong hover:bg-leaf-50 bg-white px-4 py-2 rounded font-semibold cursor-pointer"
                        >
                          취소
                        </button>
                        <button
                          onClick={handleSaveDashboardData}
                          className="text-xs bg-leaf-600 hover:bg-leaf-700 text-white px-5 py-2 rounded font-semibold shadow-xs cursor-pointer"
                        >
                          저장하기
                        </button>
                      </div>
                    </div>
                  ) : (
                    <>
                      {/* 상단 타이틀 */}
                      <div className="flex flex-col md:flex-row md:justify-between md:items-center gap-4 border-b border-line pb-4">
                        <div className="flex items-center gap-4">
                          <h2 className="font-serif-title text-2xl font-bold text-ink">{selectedStudent.name}</h2>
                          <span className="text-ink-soft text-sm font-medium">
                            {selectedStudent.className} · {selectedStudent.number ? `${selectedStudent.number}번` : "번호 미지정"}
                          </span>
                        </div>
                        
                        <div className="flex items-center gap-4 flex-wrap">
                          <button
                            onClick={() => {
                              // Deep copy selectedResult
                              const copy = JSON.parse(JSON.stringify(selectedResult));
                              if (!copy.achievement_level) {
                                copy.achievement_level = { first: "C", second: "A" };
                              }
                              const domains = ["지식이해", "과정기능", "가치태도"] as const;
                              if (!copy.first_scores) copy.first_scores = {};
                              if (!copy.second_scores) copy.second_scores = {};
                              if (!copy.growth_analysis) copy.growth_analysis = {};
                              
                              domains.forEach((d: any) => {
                                if (!copy.first_scores[d]) {
                                  copy.first_scores[d] = { score: 0, max: 5, normalized: 0, evidence: "" };
                                }
                                if (!copy.second_scores[d]) {
                                  copy.second_scores[d] = { score: 0, max: 5, normalized: 0, evidence: "" };
                                }
                                if (!copy.growth_analysis[d]) {
                                  copy.growth_analysis[d] = "";
                                }
                              });
                              
                              setEditDashboardData(copy);
                              setIsEditingDashboard(true);
                            }}
                            className="text-xs font-bold bg-leaf-50 text-leaf-700 hover:bg-leaf-100 border border-leaf-200 px-3 py-2 rounded-md flex items-center gap-1.5 cursor-pointer no-print"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                            <span>검토 및 수정</span>
                          </button>
                        </div>
                      </div>

                      {/* 방사형 그래프 + 세부 점수 */}
                      <div className="flex flex-col md:flex-row gap-8 items-center md:items-start">
                        
                        {/* 방사형 그래프 컴포넌트 */}
                        <div className="w-[240px] flex flex-col items-center justify-center shrink-0">
                          <div className="relative w-[220px] h-[220px]">
                            {renderRadarChart(selectedResult)}
                          </div>
                          <div className="mt-4 flex gap-4 text-[11px] font-medium text-ink-soft no-print">
                            <div className="flex items-center gap-1.5">
                              <span className="w-3.5 h-3 bg-leaf-200 border border-leaf-400" />
                              <span>1차 평가</span>
                            </div>
                            <div className="flex items-center gap-1.5">
                              <span className="w-3.5 h-3 bg-leaf-600 border border-leaf-700" />
                              <span>2차 평가</span>
                            </div>
                          </div>
                        </div>

                        {/* 영역별 점수 (5점 척도 환산) 및 평가 근거 */}
                        <div className="flex-1 w-full flex flex-col gap-4">
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                            {/* 지식이해 */}
                            <div className="p-3.5 border border-line rounded-lg bg-leaf-50/20">
                              <div className="text-[10px] font-bold text-ink-soft uppercase tracking-wider mb-1">
                                지식·이해 점수 (5점 환산)
                              </div>
                              <div className="text-[18px] font-serif font-bold text-ink flex items-baseline gap-1.5">
                                {selectedResult.first_scores.지식이해.normalized !== null ? selectedResult.first_scores.지식이해.normalized.toFixed(1) : "—"}
                                <span className="text-xs text-ink-muted font-normal font-sans">→</span>
                                <span className="text-leaf-700">{selectedResult.second_scores.지식이해.normalized !== null ? selectedResult.second_scores.지식이해.normalized.toFixed(1) : "—"}</span>
                              </div>
                              {selectedResult.second_scores.지식이해.normalized !== null && selectedResult.first_scores.지식이해.normalized !== null && (
                                <div className="text-[11px] text-leaf-600 font-bold mt-1">
                                  ▲ +{(selectedResult.second_scores.지식이해.normalized - selectedResult.first_scores.지식이해.normalized).toFixed(1)} 성장
                                </div>
                              )}
                            </div>

                            {/* 과정기능 */}
                            <div className="p-3.5 border border-line rounded-lg bg-leaf-50/20">
                              <div className="text-[10px] font-bold text-ink-soft uppercase tracking-wider mb-1">
                                과정·기능 점수 (5점 환산)
                              </div>
                              <div className="text-[18px] font-serif font-bold text-ink flex items-baseline gap-1.5">
                                {selectedResult.first_scores.과정기능.normalized !== null ? selectedResult.first_scores.과정기능.normalized.toFixed(1) : "—"}
                                <span className="text-xs text-ink-muted font-normal font-sans">→</span>
                                <span className="text-leaf-700">{selectedResult.second_scores.과정기능.normalized !== null ? selectedResult.second_scores.과정기능.normalized.toFixed(1) : "—"}</span>
                              </div>
                              {selectedResult.second_scores.과정기능.normalized !== null && selectedResult.first_scores.과정기능.normalized !== null && (
                                <div className="text-[11px] text-leaf-600 font-bold mt-1">
                                  ▲ +{(selectedResult.second_scores.과정기능.normalized - selectedResult.first_scores.과정기능.normalized).toFixed(1)} 성장
                                </div>
                              )}
                            </div>

                            {/* 가치태도 */}
                            <div className="p-3.5 border border-line rounded-lg bg-leaf-50/20">
                              <div className="text-[10px] font-bold text-ink-soft uppercase tracking-wider mb-1">
                                가치·태도 점수 (5점 환산)
                              </div>
                              <div className="text-[18px] font-serif font-bold text-ink flex items-baseline gap-1.5">
                                {selectedResult.first_scores.가치태도.normalized !== null ? selectedResult.first_scores.가치태도.normalized.toFixed(1) : "—"}
                                <span className="text-xs text-ink-muted font-normal font-sans">→</span>
                                <span className="text-leaf-700">{selectedResult.second_scores.가치태도.normalized !== null ? selectedResult.second_scores.가치태도.normalized.toFixed(1) : "—"}</span>
                              </div>
                              {selectedResult.second_scores.가치태도.normalized !== null && selectedResult.first_scores.가치태도.normalized !== null && (
                                <div className="text-[11px] text-leaf-600 font-bold mt-1">
                                  ▲ +{(selectedResult.second_scores.가치태도.normalized - selectedResult.first_scores.가치태도.normalized).toFixed(1)} 성장
                                </div>
                              )}
                            </div>
                          </div>

                          {/* 성장 종합 요약 문구 */}
                          <div className="p-4 rounded-lg bg-leaf-50 border border-leaf-100 flex flex-col gap-1.5">
                            <span className="small-caps text-leaf-600 text-xs font-semibold uppercase tracking-wider">성장 종합 서평</span>
                            <p className="text-sm text-ink-soft leading-relaxed font-medium">
                              {selectedResult.overall_summary}
                            </p>
                          </div>
                        </div>

                      </div>

                      {/* 영역별 세부 변화 기술서 */}
                      <div className="flex flex-col gap-4 mt-2">
                        <h4 className="font-serif text-base font-bold text-ink border-b border-line pb-2 flex items-center gap-2">
                          영역별 학습 성장 정성 분석
                        </h4>
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                          <div className="flex flex-col gap-1.5">
                            <span className="text-xs font-bold text-leaf-600 flex items-center gap-1.5">
                              <span className="w-1.5 h-1.5 bg-leaf-600 rounded-full" />
                              지식·이해 변화
                            </span>
                            <p className="text-xs text-ink-soft leading-relaxed bg-[#faf9f4] p-3 rounded border border-line">
                              {selectedResult.growth_analysis.지식이해}
                            </p>
                            <span className="text-[10px] text-ink-muted leading-snug">
                              * 채점 근거: {selectedResult.second_scores.지식이해.evidence}
                            </span>
                          </div>

                          <div className="flex flex-col gap-1.5">
                            <span className="text-xs font-bold text-leaf-600 flex items-center gap-1.5">
                              <span className="w-1.5 h-1.5 bg-leaf-600 rounded-full" />
                              과정·기능 변화
                            </span>
                            <p className="text-xs text-ink-soft leading-relaxed bg-[#faf9f4] p-3 rounded border border-line">
                              {selectedResult.growth_analysis.과정기능}
                            </p>
                            <span className="text-[10px] text-ink-muted leading-snug">
                              * 채점 근거: {selectedResult.second_scores.과정기능.evidence}
                            </span>
                          </div>

                          <div className="flex flex-col gap-1.5">
                            <span className="text-xs font-bold text-leaf-600 flex items-center gap-1.5">
                              <span className="w-1.5 h-1.5 bg-leaf-600 rounded-full" />
                              가치·태도 변화
                            </span>
                            <p className="text-xs text-ink-soft leading-relaxed bg-[#faf9f4] p-3 rounded border border-line">
                              {selectedResult.growth_analysis.가치태도}
                            </p>
                            <span className="text-[10px] text-ink-muted leading-snug">
                              * * 채점 근거: {selectedResult.second_scores.가치태도.evidence}
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* 다음 단계 지도 제안 (Amber 배너) */}
                      <div className="bg-amber-soft p-4 border-l-4 border-amber rounded-r-lg flex flex-col gap-1.5">
                        <span className="small-caps text-amber text-xs font-semibold uppercase tracking-wider">
                          개인별 맞춤 다음 단계 학습 지도 전략
                        </span>
                        <p className="text-sm text-amber font-medium leading-relaxed italic">
                          &ldquo;{selectedResult.teaching_feedback}&rdquo;
                        </p>
                      </div>
                    </>
                  )
                ) : null}

              </div>

            </div>

          </div>
        )}

      </main>

      {/* --- Footer (No-Print) --- */}
      <footer className="mt-auto py-6 border-t border-line text-center text-xs text-ink-muted no-print">
        <p>© 2026 송운초 서·논술형 AI기반 평가 분석. All Rights Reserved.</p>
      </footer>

      {/* ==================== 학생 추가 등록 모달 ==================== */}
      {showAddModal && (
        <div className="fixed inset-0 z-[99999] bg-black/45 backdrop-blur-xs flex items-center justify-center p-4 no-print animate-fade-in">
          <div className="bg-white border border-line rounded-lg max-w-md w-full p-6 shadow-xl relative animate-scale-up">
            <button
              onClick={() => {
                setNewStudent({ name: "", className: "", number: "" });
                setAddMethod("single");
                setShowAddModal(false);
              }}
              className="absolute top-4 right-4 text-ink-muted hover:text-ink cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
            <h3 className="font-serif text-lg font-bold text-ink mb-4">
              새로운 학생 추가 등록
            </h3>

            {/* 등록 방법 탭 토글 */}
            <div className="flex bg-leaf-50 p-1 rounded-full border border-leaf-100 mb-5">
              <button
                type="button"
                onClick={() => setAddMethod("single")}
                className={`flex-1 py-1.5 rounded-full text-xs font-semibold text-center transition-all cursor-pointer ${
                  addMethod === "single"
                    ? "bg-leaf-600 text-white shadow-xs"
                    : "text-ink-muted hover:text-ink-soft"
                }`}
              >
                개별 직접 등록
              </button>
              <button
                type="button"
                onClick={() => setAddMethod("excel")}
                className={`flex-1 py-1.5 rounded-full text-xs font-semibold text-center transition-all cursor-pointer ${
                  addMethod === "excel"
                    ? "bg-leaf-600 text-white shadow-xs"
                    : "text-ink-muted hover:text-ink-soft"
                }`}
              >
                엑셀 파일 일괄 등록
              </button>
            </div>
            
            {addMethod === "single" ? (
              <form onSubmit={handleAddStudentSubmit} className="flex flex-col gap-4">
                <div className="flex flex-col gap-1.5">
                  <label className="text-xs font-semibold text-ink-soft">이름 (필수)</label>
                  <input
                    type="text"
                    required
                    value={newStudent.name}
                    onChange={(e) => setNewStudent({ ...newStudent, name: e.target.value })}
                    placeholder="예: 박성은"
                    className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                  />
                </div>

                <div className="grid grid-cols-3 gap-2.5">
                  <div className="flex flex-col gap-1.5">
                    <label className="text-xs font-semibold text-ink-soft">학년</label>
                    <input
                      type="text"
                      value={newStudent.grade}
                      onChange={(e) => setNewStudent({ ...newStudent, grade: e.target.value })}
                      placeholder="예: 4"
                      className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                    />
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <label className="text-xs font-semibold text-ink-soft">반</label>
                    <input
                      type="text"
                      value={newStudent.classVal}
                      onChange={(e) => setNewStudent({ ...newStudent, classVal: e.target.value })}
                      placeholder="예: 3"
                      className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                    />
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <label className="text-xs font-semibold text-ink-soft">번호</label>
                    <input
                      type="text"
                      value={newStudent.number}
                      onChange={(e) => setNewStudent({ ...newStudent, number: e.target.value })}
                      placeholder="예: 15"
                      className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                    />
                  </div>
                </div>

                <div className="flex justify-end gap-2 mt-2">
                  <button
                    type="button"
                    onClick={() => {
                      setNewStudent({ name: "", grade: "", classVal: "", number: "" });
                      setShowAddModal(false);
                    }}
                    className="text-xs border border-line-strong px-4 py-2 rounded font-medium hover:bg-leaf-50 cursor-pointer"
                  >
                    취소
                  </button>
                  <button
                    type="submit"
                    className="text-xs bg-leaf-600 text-white px-4 py-2 rounded font-medium hover:opacity-95 cursor-pointer"
                  >
                    학생 등록
                  </button>
                </div>
              </form>
            ) : (
              <div className="flex flex-col gap-4">
                <div className="bg-leaf-50/50 rounded-lg p-3 border border-leaf-100 text-xs text-ink-soft">
                  <span className="font-semibold text-leaf-700 block mb-1.5">💡 권장 엑셀 구성 형식</span>
                  <p className="mb-2 leading-relaxed">
                    첫 번째 행(A1~D1)에 아래 이름의 헤더를 입력한 뒤, 그 아래 행부터 학생 데이터를 채워주세요.
                  </p>
                  
                  {/* 예시 엑셀 표 스타일 */}
                  <div className="border border-line rounded bg-white overflow-hidden font-mono text-[10px] text-center">
                    <div className="grid grid-cols-4 bg-leaf-50 border-b border-line font-semibold text-leaf-800">
                      <div className="py-1 border-r border-line">A</div>
                      <div className="py-1 border-r border-line">B</div>
                      <div className="py-1 border-r border-line">C</div>
                      <div className="py-1">D</div>
                    </div>
                    <div className="grid grid-cols-4 border-b border-line bg-[#fafafa] font-semibold text-ink-muted">
                      <div className="py-1 border-r border-line">학년</div>
                      <div className="py-1 border-r border-line">반</div>
                      <div className="py-1 border-r border-line">번호</div>
                      <div className="py-1">이름</div>
                    </div>
                    <div className="grid grid-cols-4 text-ink-soft">
                      <div className="py-1 border-r border-line">4</div>
                      <div className="py-1 border-r border-line">3</div>
                      <div className="py-1 border-r border-line">15</div>
                      <div className="py-1">홍길동</div>
                    </div>
                  </div>
                </div>

                {/* 엑셀 업로드 파일 업로더 */}
                <div className="relative border-2 border-dashed border-leaf-200 rounded-lg p-6 bg-leaf-50/20 text-center flex flex-col items-center justify-center hover:bg-leaf-50/40 transition-colors">
                  <Upload className="w-8 h-8 text-leaf-600 mb-2" />
                  <span className="text-xs font-semibold text-ink mb-1">
                    여기를 클릭하여 엑셀 파일을 선택하세요
                  </span>
                  <span className="text-[10px] text-ink-muted">
                    지원 포맷: .xlsx, .xls
                  </span>
                  
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={handleExcelUpload}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    style={{ position: 'absolute', top: 0, left: 0, opacity: 0 }}
                  />
                </div>

                <div className="flex justify-end gap-2 mt-1">
                  <button
                    type="button"
                    onClick={() => {
                      setAddMethod("single");
                      setShowAddModal(false);
                    }}
                    className="text-xs border border-line-strong px-4 py-2 rounded font-medium hover:bg-leaf-50 cursor-pointer"
                  >
                    닫기
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ==================== 학생 정보 수정 모달 ==================== */}
      {editingStudent && (
        <div className="fixed inset-0 z-[99999] bg-black/45 backdrop-blur-xs flex items-center justify-center p-4 no-print animate-fade-in">
          <div className="bg-white border border-line rounded-lg max-w-md w-full p-6 shadow-xl relative animate-scale-up">
            <button
              onClick={() => setEditingStudent(null)}
              className="absolute top-4 right-4 text-ink-muted hover:text-ink cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
            <h3 className="font-serif text-lg font-bold text-ink mb-4">
              학생 기본 정보 변경
            </h3>
            
            <form onSubmit={handleEditStudentSubmit} className="flex flex-col gap-4">
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-semibold text-ink-soft">이름 (필수)</label>
                <input
                  type="text"
                  required
                  value={editingStudent.name}
                  onChange={(e) => setEditingStudent({ ...editingStudent, name: e.target.value })}
                  placeholder="예: 박성은"
                  className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div className="flex flex-col gap-1.5">
                  <label className="text-xs font-semibold text-ink-soft">학반</label>
                  <input
                    type="text"
                    value={editingStudent.className}
                    onChange={(e) => setEditingStudent({ ...editingStudent, className: e.target.value })}
                    placeholder="예: 4반"
                    className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-xs font-semibold text-ink-soft">번호</label>
                  <input
                    type="text"
                    value={editingStudent.number}
                    onChange={(e) => setEditingStudent({ ...editingStudent, number: e.target.value })}
                    placeholder="예: 2"
                    className="w-full px-3 py-2 border border-line-strong rounded-md text-sm focus:outline-none focus:ring-1 focus:ring-leaf-400 bg-[#fafafa]"
                  />
                </div>
              </div>

              <div className="flex justify-end gap-2 mt-2">
                <button
                  type="button"
                  onClick={() => setEditingStudent(null)}
                  className="text-xs border border-line-strong px-4 py-2 rounded font-medium hover:bg-leaf-50 cursor-pointer"
                >
                  취소
                </button>
                <button
                  type="submit"
                  className="text-xs bg-leaf-600 text-white px-4 py-2 rounded font-medium hover:opacity-95 cursor-pointer"
                >
                  변경 사항 저장
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ==================== 전체 삭제 확인 경고 모달 ==================== */}
      {showClearConfirmModal && (
        <div className="fixed inset-0 z-[99999] bg-black/45 backdrop-blur-xs flex items-center justify-center p-4 no-print animate-fade-in">
          <div className="bg-white border border-coral rounded-lg max-w-sm w-full p-6 shadow-xl relative animate-scale-up">
            <h3 className="font-serif text-lg font-bold text-coral mb-2 flex items-center gap-2">
              <AlertCircle className="w-5 h-5 shrink-0" />
              학급 데이터 전체 삭제
            </h3>
            <p className="text-xs text-ink-soft leading-relaxed mb-4">
              이 작업은 되돌릴 수 없습니다. IndexedDB 스토리지 내 등록된 모든 루브릭 가이드 파일, 
              학생 명단, 업로드된 1차/2차 평가지 이미지, 성 분석 기록이 완전히 포맷 및 말소됩니다. 
              정말로 계속하시겠습니까?
            </p>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setShowClearConfirmModal(false)}
                className="text-xs border border-line-strong px-4 py-2 rounded font-medium hover:bg-leaf-50 cursor-pointer"
              >
                아니오, 보존합니다
              </button>
              <button
                onClick={handleClearAllData}
                className="text-xs bg-coral text-white px-4 py-2 rounded font-medium hover:opacity-90 cursor-pointer"
              >
                예, 전체 파기합니다
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== 학생 개별 삭제 확인 모달 ==================== */}
      {studentToDelete && (
        <div className="fixed inset-0 z-[99999] bg-black/45 backdrop-blur-xs flex items-center justify-center p-4 no-print animate-fade-in">
          <div className="bg-white border border-line rounded-lg max-w-sm w-full p-6 shadow-xl relative animate-scale-up">
            <h3 className="font-serif text-lg font-bold text-coral mb-2 flex items-center gap-2">
              <AlertCircle className="w-5 h-5 shrink-0" />
              학생 데이터 삭제
            </h3>
            <p className="text-xs text-ink-soft leading-relaxed mb-4">
              <strong>{studentToDelete.name}</strong> 학생의 모든 평가 데이터, 업로드된 평가지 이미지 및 AI 분석 결과가 완전히 삭제됩니다. 이 작업은 되돌릴 수 없습니다. 정말 삭제하시겠습니까?
            </p>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setStudentToDelete(null)}
                className="text-xs border border-line-strong px-4 py-2 rounded font-medium hover:bg-leaf-50 cursor-pointer"
              >
                취소
              </button>
              <button
                onClick={confirmDeleteStudent}
                className="text-xs bg-coral text-white px-4 py-2 rounded font-medium hover:opacity-90 cursor-pointer"
              >
                삭제하기
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
