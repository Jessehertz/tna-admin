"use client";

import Breadcrumb from "@/components/Breadcrumbs/Breadcrumb";
import InputGroup from "@/components/FormElements/InputGroup";
import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

interface Question {
  id: string;
  question: string;
}

interface ModuleData {
  id?: string;
  moduleName: string;
  moduleDescription: string;
  category: string;
  questions: Question[];
  status?: string;
  dateCreated?: string;
}

type ManageModulesViewProps = {
  searchParams?: Promise<{ id?: string }>;
};

export default function ManageModulesViewPage({
  searchParams,
}: ManageModulesViewProps) {
  const [moduleName, setModuleName] = useState("");
  const [moduleDescription, setModuleDescription] = useState("");
  const [category, setCategory] = useState("");
  const [status, setStatus] = useState<string>("active");
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [previewRatings, setPreviewRatings] = useState<Record<string, number>>(
    {},
  );
  const [isEditing, setIsEditing] = useState(false);
  const [moduleId, setModuleId] = useState<string | null>(null);
  const [questions, setQuestions] = useState<Question[]>([
    {
      id: "1",
      question: "",
    },
  ]);

  // Load module data (this would typically come from an API)
  useEffect(() => {
    // In a real app, you would fetch module data based on searchParams.id
    // For now, this is a placeholder
    const loadModuleData = async () => {
      // Example: const module = await fetchModuleById(id);
      // setModuleName(module.name);
      // setModuleDescription(module.description);
      // setCategory(module.category);
      // setQuestions(module.questions);
      // setModuleId(module.id);
    };

    if (searchParams) {
      searchParams.then((params) => {
        if (params.id) {
          setModuleId(params.id);
          loadModuleData();
        }
      });
    }
  }, [searchParams]);

  const addQuestion = () => {
    setQuestions([
      ...questions,
      {
        id: Date.now().toString(),
        question: "",
      },
    ]);
  };

  const removeQuestion = (id: string) => {
    setQuestions(questions.filter((q) => q.id !== id));
  };

  const updateQuestion = (id: string, field: keyof Question, value: any) => {
    setQuestions(
      questions.map((q) => (q.id === id ? { ...q, [field]: value } : q)),
    );
  };

  const handleExcelUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setExcelFile(file);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });

      // Get the first sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Convert to JSON
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Parse module information (assuming first row has headers, second row has data)
      if (data.length > 1) {
        const headers = data[0] as string[];
        const moduleData = data[1] as any[];

        // Find column indices
        const nameIndex = headers.findIndex(
          (h) =>
            h &&
            (h.toString().toLowerCase().includes("module name") ||
              h.toString().toLowerCase().includes("name")),
        );
        const descIndex = headers.findIndex(
          (h) =>
            h &&
            (h.toString().toLowerCase().includes("description") ||
              h.toString().toLowerCase().includes("desc")),
        );
        const categoryIndex = headers.findIndex(
          (h) =>
            h &&
            (h.toString().toLowerCase().includes("category") ||
              h.toString().toLowerCase().includes("cat")),
        );

        // Set module information
        if (nameIndex >= 0 && moduleData[nameIndex]) {
          setModuleName(String(moduleData[nameIndex]));
        }
        if (descIndex >= 0 && moduleData[descIndex]) {
          setModuleDescription(String(moduleData[descIndex]));
        }
        if (categoryIndex >= 0 && moduleData[categoryIndex]) {
          setCategory(String(moduleData[categoryIndex]).toLowerCase());
        }
      }

      // Parse questions (look for a "Questions" sheet or a column with questions)
      let questionsSheet = workbook.Sheets[firstSheetName];
      const questionsSheetName = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes("question"),
      );
      if (questionsSheetName) {
        questionsSheet = workbook.Sheets[questionsSheetName];
      }

      const questionsData = XLSX.utils.sheet_to_json(questionsSheet, {
        header: 1,
      });

      // Find questions column
      if (questionsData.length > 0) {
        const questionHeaders = questionsData[0] as string[];
        const questionColIndex = questionHeaders.findIndex(
          (h) =>
            h &&
            (h.toString().toLowerCase().includes("question") ||
              h.toString().toLowerCase().includes("assessment")),
        );

        if (questionColIndex >= 0) {
          // Extract questions (skip header row)
          const extractedQuestions: Question[] = [];
          for (let i = 1; i < questionsData.length; i++) {
            const row = questionsData[i] as any[];
            const questionText = row[questionColIndex];
            if (questionText && String(questionText).trim()) {
              extractedQuestions.push({
                id: Date.now().toString() + i,
                question: String(questionText).trim(),
              });
            }
          }

          if (extractedQuestions.length > 0) {
            setQuestions(extractedQuestions);
          }
        } else {
          // If no question column found, try to find questions in any column
          // Look for rows that might be questions (non-empty cells)
          const extractedQuestions: Question[] = [];
          for (let i = 1; i < questionsData.length; i++) {
            const row = questionsData[i] as any[];
            // Check each cell in the row
            for (let j = 0; j < row.length; j++) {
              const cellValue = row[j];
              if (
                cellValue &&
                String(cellValue).trim() &&
                String(cellValue).length > 10 // Likely a question if it's longer
              ) {
                extractedQuestions.push({
                  id: Date.now().toString() + i + j,
                  question: String(cellValue).trim(),
                });
                break; // Take first valid question from each row
              }
            }
          }
          if (extractedQuestions.length > 0) {
            setQuestions(extractedQuestions);
          }
        }
      }

      alert("Excel file parsed successfully! Form fields have been populated.");
    } catch (error) {
      console.error("Error parsing Excel file:", error);
      alert(
        "Error parsing Excel file. Please ensure the file format is correct.\n\nExpected format:\n- First row: Headers (Module Name, Description, Category, Questions)\n- Subsequent rows: Data",
      );
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    // Handle form submission
    const moduleData: ModuleData = {
      id: moduleId || undefined,
      moduleName,
      moduleDescription,
      category,
      status: status || "active",
      questions,
    };

    console.log(moduleData);

    if (moduleId) {
      // Update existing module
      alert("Module updated successfully!");
    } else {
      // Create new module
      alert("Module created successfully!");
    }
  };

  const handleDelete = () => {
    if (confirm("Are you sure you want to delete this module?")) {
      // Handle deletion
      console.log("Deleting module:", moduleId);
      alert("Module deleted successfully!");
    }
  };

  return (
    <>
      <Breadcrumb pageName={moduleId ? "Edit Module" : "View Module"} />

      <div className="rounded-[10px] bg-white p-6.5 shadow-1 dark:bg-gray-dark dark:shadow-card">
        <form onSubmit={handleSubmit} className="space-y-6">
          {/* Module Header */}
          <div className="mb-6 flex items-center justify-between">
            <div>
              {moduleId && (
                <p className="mt-1 text-sm text-dark-6 dark:text-dark-4">
                  Module ID: {moduleId}
                </p>
              )}
            </div>
            <div className="flex gap-2">
              {moduleId && !isEditing && (
                <button
                  type="button"
                  onClick={() => setIsEditing(true)}
                  className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-6 py-2 text-center text-sm font-medium text-primary transition hover:bg-primary/10 focus:outline-none"
                >
                  Edit Module
                </button>
              )}
              {isEditing && (
                <button
                  type="button"
                  onClick={() => setIsEditing(false)}
                  className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-stroke px-6 py-2 text-center text-sm font-medium text-dark transition hover:bg-gray-100 focus:outline-none dark:border-dark-3 dark:text-white dark:hover:bg-dark-2"
                >
                  Cancel Edit
                </button>
              )}
              {moduleId && (
                <button
                  type="button"
                  onClick={handleDelete}
                  className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-red px-6 py-2 text-center text-sm font-medium text-red transition hover:bg-red/10 focus:outline-none"
                >
                  Delete
                </button>
              )}
            </div>
          </div>

          {/* Excel Upload Section - Only show when editing */}
          {isEditing && (
            <div className="rounded-lg border-2 border-dashed border-stroke p-6 dark:border-dark-3">
              <div className="text-center">
                <h3 className="mb-2 text-lg font-semibold text-dark dark:text-white">
                  Upload Excel File (Optional)
                </h3>
                <p className="mb-4 text-sm text-dark-6 dark:text-dark-4">
                  Upload an Excel file to automatically populate module
                  information and assessment questions
                </p>
                <div className="mx-auto max-w-2xl rounded-lg bg-gray-100 p-4 text-left dark:bg-dark-2">
                  <p className="mb-2 text-xs font-semibold text-dark dark:text-white">
                    Excel Format:
                  </p>
                  <p className="mb-2 text-xs text-dark-6 dark:text-dark-4">
                    <strong>Option 1 - Single Sheet:</strong>
                  </p>
                  <div className="mb-3 overflow-x-auto rounded border border-stroke bg-white p-2 text-xs dark:border-dark-3 dark:bg-gray-dark">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="border-b border-stroke dark:border-dark-3">
                          <th className="p-1 text-left">Module Name</th>
                          <th className="p-1 text-left">Description</th>
                          <th className="p-1 text-left">Category</th>
                          <th className="p-1 text-left">Question</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td className="p-1">Tax Basics</td>
                          <td className="p-1">Introduction to taxation</td>
                          <td className="p-1">taxation</td>
                          <td className="p-1">
                            How well do you understand tax compliance?
                          </td>
                        </tr>
                        <tr>
                          <td className="p-1"></td>
                          <td className="p-1"></td>
                          <td className="p-1"></td>
                          <td className="p-1">
                            Rate your knowledge of tax deductions
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <p className="mb-2 text-xs text-dark-6 dark:text-dark-4">
                    <strong>Option 2 - Separate Questions Sheet:</strong>
                  </p>
                  <p className="text-xs text-dark-6 dark:text-dark-4">
                    Create a sheet named "Questions" with a "Question" column
                    header. Each row below is a question.
                  </p>
                </div>
                <InputGroup
                  type="file"
                  fileStyleVariant="style1"
                  label=""
                  placeholder="Choose Excel file (.xlsx, .xls)"
                  handleChange={handleExcelUpload}
                  className="mx-auto max-w-md"
                />
                {excelFile && (
                  <p className="mt-2 text-sm text-green">
                    File selected: {excelFile.name}
                  </p>
                )}
              </div>
            </div>
          )}

          {/* Module Basic Information */}
          <div className="space-y-5">
            <h3 className="text-lg font-semibold text-dark dark:text-white">
              Module Information
            </h3>

            <div className="grid gap-5 sm:grid-cols-2">
              <InputGroup
                label="Module Name"
                placeholder="Enter module name"
                type="text"
                required
                value={moduleName}
                handleChange={(e) => setModuleName(e.target.value)}
                disabled={!isEditing && !!moduleId}
              />

              <div className="space-y-3">
                <label className="block text-body-sm font-medium text-dark dark:text-white">
                  Category
                </label>
                <select
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  disabled={!isEditing && !!moduleId}
                  className="w-full appearance-none rounded-lg border border-stroke bg-transparent px-5.5 py-3 text-dark outline-none transition focus:border-primary active:border-primary disabled:cursor-not-allowed disabled:bg-gray-2 disabled:opacity-50 dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary dark:disabled:bg-dark [&>option]:text-dark-5 dark:[&>option]:text-dark-6"
                >
                  <option value="" disabled>
                    Select category
                  </option>
                  <option value="taxation">Taxation</option>
                  <option value="compliance">Compliance</option>
                  <option value="accounting">Accounting</option>
                  <option value="general">General</option>
                </select>
              </div>
            </div>

            <div className="space-y-3">
              <label className="block text-body-sm font-medium text-dark dark:text-white">
                Module Description
              </label>
              <textarea
                value={moduleDescription}
                onChange={(e) => setModuleDescription(e.target.value)}
                rows={4}
                placeholder="Enter module description"
                disabled={!isEditing && !!moduleId}
                className="w-full rounded-lg border-[1.5px] border-stroke bg-transparent px-5.5 py-3 text-dark outline-none transition focus:border-primary disabled:cursor-not-allowed disabled:bg-gray-2 disabled:opacity-50 dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary dark:disabled:bg-dark"
              />
            </div>

            <div className="space-y-3">
              <label className="block text-body-sm font-medium text-dark dark:text-white">
                Status
              </label>
              <div className="flex gap-6">
                <label className="flex cursor-pointer items-center">
                  <input
                    type="radio"
                    name="status"
                    value="active"
                    checked={status === "active"}
                    onChange={(e) => setStatus(e.target.value)}
                    disabled={!isEditing && !!moduleId}
                    className="mr-2 h-4 w-4 cursor-pointer border-stroke text-primary focus:ring-2 focus:ring-primary disabled:cursor-not-allowed disabled:opacity-50 dark:border-dark-3"
                  />
                  <span className="text-body-sm text-dark dark:text-white">
                    Active
                  </span>
                </label>
                <label className="flex cursor-pointer items-center">
                  <input
                    type="radio"
                    name="status"
                    value="inactive"
                    checked={status === "inactive"}
                    onChange={(e) => setStatus(e.target.value)}
                    disabled={!isEditing && !!moduleId}
                    className="mr-2 h-4 w-4 cursor-pointer border-stroke text-primary focus:ring-2 focus:ring-primary disabled:cursor-not-allowed disabled:opacity-50 dark:border-dark-3"
                  />
                  <span className="text-body-sm text-dark dark:text-white">
                    Inactive
                  </span>
                </label>
              </div>
            </div>
          </div>

          {/* Assessment Questions Section */}
          <div className="space-y-5">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-lg font-semibold text-dark dark:text-white">
                  Assessment Questions
                </h3>
                <p className="mt-1 text-sm text-dark-6 dark:text-dark-4">
                  Each question will be rated 1-5 by staff members to assess
                  their knowledge level
                </p>
              </div>
              {(isEditing || !moduleId) && (
                <button
                  type="button"
                  onClick={addQuestion}
                  className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-6 py-[11px] text-center font-medium text-primary transition hover:bg-primary/10 focus:outline-none"
                >
                  Add Assessment Question
                </button>
              )}
            </div>

            {questions.map((question, index) => (
              <div
                key={question.id}
                className="rounded-lg border border-stroke p-5 dark:border-dark-3"
              >
                <div className="mb-4 flex items-center justify-between">
                  <h4 className="font-semibold text-dark dark:text-white">
                    Assessment Question {index + 1}
                  </h4>
                  {questions.length > 1 && (
                    <button
                      type="button"
                      onClick={() => removeQuestion(question.id)}
                      className="text-sm text-red hover:text-red/80"
                    >
                      Remove
                    </button>
                  )}
                </div>

                <div className="space-y-4">
                  <div className="space-y-3">
                    <label className="block text-body-sm font-medium text-dark dark:text-white">
                      Question Text
                    </label>
                    <textarea
                      value={question.question}
                      onChange={(e) =>
                        updateQuestion(question.id, "question", e.target.value)
                      }
                      rows={3}
                      placeholder="e.g., How well do you understand tax compliance requirements?"
                      disabled={!isEditing && !!moduleId}
                      className="w-full rounded-lg border-[1.5px] border-stroke bg-transparent px-5.5 py-3 text-dark outline-none transition focus:border-primary disabled:cursor-not-allowed disabled:bg-gray-2 disabled:opacity-50 dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary dark:disabled:bg-dark"
                    />
                    <p className="text-xs text-dark-6 dark:text-dark-4">
                      Staff members will rate their knowledge of this topic on a
                      scale of 1-5 (1 = Poor, 5 = Excellent)
                    </p>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* Submit Buttons */}
          <div className="flex justify-end gap-4 pt-4">
            <button
              type="button"
              onClick={() => setShowPreview(!showPreview)}
              className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-10 py-3.5 text-center font-medium text-primary transition hover:bg-primary/10 focus:outline-none lg:px-8 xl:px-10"
            >
              {showPreview ? "Hide Preview" : "Preview"}
            </button>
            <button
              type="button"
              className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-10 py-3.5 text-center font-medium text-primary transition hover:bg-primary/10 focus:outline-none lg:px-8 xl:px-10"
            >
              Cancel
            </button>
            {(isEditing || !moduleId) && (
              <button
                type="submit"
                className="inline-flex items-center justify-center gap-2.5 rounded-[5px] bg-[rgb(214,49,41)] px-10 py-3.5 text-center font-medium text-white transition hover:bg-[rgb(214,49,41)]/90 focus:outline-none lg:px-8 xl:px-10"
              >
                {moduleId ? "Update Module" : "Create Module"}
              </button>
            )}
          </div>
        </form>

        {/* Preview Section */}
        {showPreview && (
          <div className="mt-8 rounded-[10px] border-2 border-stroke bg-gray-50 p-6.5 dark:border-dark-3 dark:bg-dark-2">
            <div className="mb-6 flex items-center justify-between">
              <h2 className="text-body-2xlg font-bold text-dark dark:text-white">
                Assessment Preview
              </h2>
              <span className="rounded-full bg-green/10 px-3 py-1 text-xs font-medium text-green">
                Staff View
              </span>
            </div>

            {/* Module Info Preview */}
            <div className="mb-6 rounded-lg border border-stroke bg-white p-5 dark:border-dark-3 dark:bg-gray-dark">
              <h3 className="mb-2 text-xl font-bold text-dark dark:text-white">
                {moduleName || "Module Name"}
              </h3>
              {category && (
                <span className="mb-3 inline-block rounded-full bg-primary/10 px-3 py-1 text-xs font-medium text-primary">
                  {category.charAt(0).toUpperCase() + category.slice(1)}
                </span>
              )}
              {moduleDescription && (
                <p className="mt-3 text-sm text-dark-6 dark:text-dark-4">
                  {moduleDescription}
                </p>
              )}
            </div>

            {/* Questions Preview */}
            <div className="space-y-4">
              <h3 className="text-lg font-semibold text-dark dark:text-white">
                Rate your knowledge level for each topic (1 = Poor, 5 =
                Excellent)
              </h3>

              {questions.map((question, index) => {
                if (!question.question.trim()) return null;

                return (
                  <div
                    key={question.id}
                    className="rounded-lg border border-stroke bg-white p-5 dark:border-dark-3 dark:bg-gray-dark"
                  >
                    <p className="mb-4 font-medium text-dark dark:text-white">
                      {index + 1}. {question.question}
                    </p>

                    <div className="flex items-center gap-4">
                      <span className="text-sm font-medium text-dark-6 dark:text-dark-4">
                        Poor
                      </span>
                      <div className="flex gap-2">
                        {[1, 2, 3, 4, 5].map((rating) => (
                          <button
                            key={rating}
                            type="button"
                            onClick={() =>
                              setPreviewRatings({
                                ...previewRatings,
                                [question.id]: rating,
                              })
                            }
                            className={`flex h-12 w-12 items-center justify-center rounded-lg border-2 font-semibold transition ${
                              previewRatings[question.id] === rating
                                ? "border-[rgb(214,49,41)] bg-[rgb(214,49,41)] text-white"
                                : "border-stroke bg-transparent text-dark hover:border-primary dark:border-dark-3 dark:text-white"
                            }`}
                          >
                            {rating}
                          </button>
                        ))}
                      </div>
                      <span className="text-sm font-medium text-dark-6 dark:text-dark-4">
                        Excellent
                      </span>
                    </div>
                  </div>
                );
              })}

              {questions.filter((q) => q.question.trim()).length === 0 && (
                <div className="rounded-lg border border-stroke bg-white p-8 text-center dark:border-dark-3 dark:bg-gray-dark">
                  <p className="text-dark-6 dark:text-dark-4">
                    Add questions to see the preview
                  </p>
                </div>
              )}

              {questions.filter((q) => q.question.trim()).length > 0 && (
                <div className="mt-6 flex justify-end">
                  <button
                    type="button"
                    className="inline-flex items-center justify-center gap-2.5 rounded-[5px] bg-[rgb(214,49,41)] px-10 py-3.5 text-center font-medium text-white transition hover:bg-[rgb(214,49,41)]/90 focus:outline-none lg:px-8 xl:px-10"
                  >
                    Submit Assessment
                  </button>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </>
  );
}
