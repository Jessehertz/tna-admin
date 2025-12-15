"use client";

import Breadcrumb from "@/components/Breadcrumbs/Breadcrumb";
import InputGroup from "@/components/FormElements/InputGroup";
import { useState } from "react";
import * as XLSX from "xlsx";

interface Question {
  id: string;
  question: string;
}

interface ModuleData {
  id: string;
  position?: string;
  moduleName: string;
  moduleDescription: string;
  questions: Question[];
}

export default function UploadNewModulePage() {
  const [moduleName, setModuleName] = useState("");
  const [moduleDescription, setModuleDescription] = useState("");
  const [position, setPosition] = useState("");
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [previewRatings, setPreviewRatings] = useState<Record<string, number>>(
    {},
  );
  const [questions, setQuestions] = useState<Question[]>([
    {
      id: "1",
      question: "",
    },
  ]);
  const [importedModules, setImportedModules] = useState<ModuleData[]>([]);
  const [isBulkMode, setIsBulkMode] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [expandedModules, setExpandedModules] = useState<Set<string>>(
    new Set(),
  );

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

  // Parse topics string into questions array
  const parseTopicsToQuestions = (topicsString: string): Question[] => {
    if (!topicsString || !String(topicsString).trim()) return [];

    const topics = String(topicsString)
      .split(/[,\n•\-\*]/)
      .map((t) => t.trim())
      .filter((t) => t.length > 0);

    return topics.map((topic, index) => ({
      id: Date.now().toString() + index,
      question: `How well do you understand ${topic}?`,
    }));
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

      if (data.length < 2) {
        alert("Excel file must have at least a header row and one data row.");
        return;
      }

      const headers = (data[0] as any[]).map((h) =>
        String(h || "").toLowerCase().trim(),
      );

      // Find column indices - flexible matching
      const positionIndex = headers.findIndex(
        (h) =>
          h.includes("position") ||
          h.includes("role") ||
          h.includes("job title") ||
          h.includes("title"),
      );
      const moduleTitleIndex = headers.findIndex(
        (h) =>
          h.includes("module title") ||
          h.includes("module name") ||
          h.includes("module") ||
          h.includes("name") ||
          h.includes("title"),
      );
      const topicsIndex = headers.findIndex(
        (h) =>
          h.includes("topic") ||
          h.includes("topics") ||
          h.includes("question") ||
          h.includes("assessment"),
      );
      const descIndex = headers.findIndex(
        (h) =>
          h.includes("description") ||
          h.includes("desc") ||
          h.includes("details"),
      );

      if (moduleTitleIndex < 0) {
        alert(
          "Could not find 'Module Title' or 'Module Name' column. Please ensure your Excel file has the correct headers.",
        );
        return;
      }

      // Parse all rows as modules
      const modules: ModuleData[] = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i] as any[];

        const moduleTitle = row[moduleTitleIndex]
          ? String(row[moduleTitleIndex]).trim()
          : "";
        const position = row[positionIndex]
          ? String(row[positionIndex]).trim()
          : "";
        const topics = row[topicsIndex] ? String(row[topicsIndex]).trim() : "";
        const description = row[descIndex]
          ? String(row[descIndex]).trim()
          : "";

        // Skip empty rows
        if (!moduleTitle) continue;

        // Convert topics to questions
        const questions = parseTopicsToQuestions(topics);

        modules.push({
          id: `module-${Date.now()}-${i}`,
          position: position || undefined,
          moduleName: moduleTitle,
          moduleDescription: description || `Module: ${moduleTitle}`,
          questions: questions.length > 0 ? questions : [
            {
              id: Date.now().toString() + i,
              question: `How well do you understand ${moduleTitle}?`,
            },
          ],
        });
      }

      if (modules.length === 0) {
        alert("No valid modules found in the Excel file.");
        return;
      }

      // If only one module, populate single module form
      if (modules.length === 1) {
        const module = modules[0];
        setModuleName(module.moduleName);
        setModuleDescription(module.moduleDescription);
        setPosition(module.position || "");
        setQuestions(module.questions);
        setIsBulkMode(false);
        alert("Excel file parsed successfully! Form fields have been populated.");
      } else {
        // Multiple modules - switch to bulk mode
        setImportedModules(modules);
        setIsBulkMode(true);
        // Expand all modules by default
        setExpandedModules(new Set(modules.map((m) => m.id)));
        alert(
          `Successfully imported ${modules.length} modules! Review and submit them below.`,
        );
      }
    } catch (error) {
      console.error("Error parsing Excel file:", error);
      alert(
        "Error parsing Excel file. Please ensure the file format is correct.\n\nExpected format:\n- First row: Headers (Position, Module Title, Topics)\n- Subsequent rows: Module data",
      );
    }
  };

  const updateImportedModule = (
    id: string,
    field: keyof ModuleData,
    value: any,
  ) => {
    setImportedModules(
      importedModules.map((m) =>
        m.id === id ? { ...m, [field]: value } : m,
      ),
    );
  };

  const updateImportedModuleQuestion = (
    moduleId: string,
    questionId: string,
    value: string,
  ) => {
    setImportedModules(
      importedModules.map((m) => {
        if (m.id === moduleId) {
          return {
            ...m,
            questions: m.questions.map((q) =>
              q.id === questionId ? { ...q, question: value } : q,
            ),
          };
        }
        return m;
      }),
    );
  };

  const removeImportedModule = (id: string) => {
    setImportedModules(importedModules.filter((m) => m.id !== id));
  };

  const addQuestionToModule = (moduleId: string) => {
    setImportedModules(
      importedModules.map((m) => {
        if (m.id === moduleId) {
          return {
            ...m,
            questions: [
              ...m.questions,
              {
                id: Date.now().toString(),
                question: "",
              },
            ],
          };
        }
        return m;
      }),
    );
  };

  const removeQuestionFromModule = (moduleId: string, questionId: string) => {
    setImportedModules(
      importedModules.map((m) => {
        if (m.id === moduleId) {
          return {
            ...m,
            questions: m.questions.filter((q) => q.id !== questionId),
          };
        }
        return m;
      }),
    );
  };

  const toggleModuleExpansion = (moduleId: string) => {
    setExpandedModules((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(moduleId)) {
        newSet.delete(moduleId);
      } else {
        newSet.add(moduleId);
      }
      return newSet;
    });
  };

  const expandAllModules = () => {
    setExpandedModules(new Set(importedModules.map((m) => m.id)));
  };

  const collapseAllModules = () => {
    setExpandedModules(new Set());
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (isBulkMode) {
      // Validate all modules
      const invalidModules = importedModules.filter(
        (m) => !m.moduleName.trim() || m.questions.length === 0,
      );

      if (invalidModules.length > 0) {
        alert(
          `Please ensure all modules have a name and at least one question. ${invalidModules.length} module(s) are invalid.`,
        );
        return;
      }

      setIsSubmitting(true);
      try {
        // Submit all modules
        for (const module of importedModules) {
          console.log("Submitting module:", module);
          // TODO: Replace with actual API call
          // await fetch('/api/modules', { method: 'POST', body: JSON.stringify(module) });
        }

        alert(
          `Successfully created ${importedModules.length} module(s)!`,
        );
        // Reset form
        setImportedModules([]);
        setIsBulkMode(false);
        setExcelFile(null);
      } catch (error) {
        console.error("Error submitting modules:", error);
        alert("Error submitting modules. Please try again.");
      } finally {
        setIsSubmitting(false);
      }
    } else {
      // Single module submission
      if (!moduleName.trim() || questions.length === 0) {
        alert("Please fill in all required fields.");
        return;
      }

      const moduleData = {
        moduleName,
        moduleDescription,
        position: position || undefined,
        questions,
      };

      console.log("Submitting module:", moduleData);
      // TODO: Replace with actual API call
      // await fetch('/api/modules', { method: 'POST', body: JSON.stringify(moduleData) });
      alert("Module created successfully!");
    }
  };

  return (
    <>
      <Breadcrumb pageName="Upload New Module" />

      <div className="rounded-[10px] bg-white p-6.5 shadow-1 dark:bg-gray-dark dark:shadow-card">
        <form onSubmit={handleSubmit} className="space-y-6">
          {/* Excel Upload Section */}
          <div className="rounded-lg border-2 border-dashed border-stroke p-6 dark:border-dark-3">
            <div className="text-center">
              <h3 className="mb-2 text-lg font-semibold text-dark dark:text-white">
                Upload Excel File for Bulk Import
              </h3>
              <p className="mb-4 text-sm text-dark-6 dark:text-dark-4">
                Upload an Excel file to import multiple modules at once. Each row
                will be imported as a separate module.
              </p>
              <div className="mx-auto max-w-2xl rounded-lg bg-gray-100 p-4 text-left dark:bg-dark-2">
                <p className="mb-2 text-xs font-semibold text-dark dark:text-white">
                  Excel Format (Bulk Import):
                </p>
                <div className="mb-3 overflow-x-auto rounded border border-stroke bg-white p-2 text-xs dark:border-dark-3 dark:bg-gray-dark">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="border-b border-stroke dark:border-dark-3">
                        <th className="p-1 text-left">Position</th>
                        <th className="p-1 text-left">Module Title</th>
                        <th className="p-1 text-left">Topics</th>
                        <th className="p-1 text-left">Description (Optional)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="p-1">Enterprise Portfolio Manager</td>
                        <td className="p-1">
                          Enterprise Portfolio Management Methodologies (MoP,
                          PfMP)
                        </td>
                        <td className="p-1">
                          Portfolio lifecycle and governance frameworks,
                          Portfolio categorization and structuring, Strategic
                          alignment and value delivery
                        </td>
                        <td className="p-1">Module description here</td>
                      </tr>
                      <tr>
                        <td className="p-1">Project Manager</td>
                        <td className="p-1">Project Management Fundamentals</td>
                        <td className="p-1">
                          Project planning, Risk management, Stakeholder
                          engagement
                        </td>
                        <td className="p-1">Module description here</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <p className="text-xs text-dark-6 dark:text-dark-4">
                  <strong>Note:</strong> Topics can be separated by commas,
                  newlines, or bullet points. Each topic will be converted to an
                  assessment question.
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

          {/* Bulk Import Preview */}
          {isBulkMode && importedModules.length > 0 && (
            <div className="space-y-5">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-lg font-semibold text-dark dark:text-white">
                    Imported Modules ({importedModules.length})
                  </h3>
                  <p className="mt-1 text-sm text-dark-6 dark:text-dark-4">
                    Review and edit the imported modules before submitting
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={expandAllModules}
                    className="inline-flex items-center justify-center gap-2 rounded-[5px] border border-primary px-4 py-2 text-sm font-medium text-primary transition hover:bg-primary/10 focus:outline-none"
                  >
                    Expand All
                  </button>
                  <button
                    type="button"
                    onClick={collapseAllModules}
                    className="inline-flex items-center justify-center gap-2 rounded-[5px] border border-primary px-4 py-2 text-sm font-medium text-primary transition hover:bg-primary/10 focus:outline-none"
                  >
                    Collapse All
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setImportedModules([]);
                      setIsBulkMode(false);
                      setExcelFile(null);
                      setExpandedModules(new Set());
                    }}
                    className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-red px-6 py-[11px] text-center font-medium text-red transition hover:bg-red/10 focus:outline-none"
                  >
                    Clear All
                  </button>
                </div>
              </div>

              <div className="space-y-3">
                {importedModules.map((module, moduleIndex) => {
                  const isExpanded = expandedModules.has(module.id);
                  return (
                    <div
                      key={module.id}
                      className="rounded-lg border-2 border-stroke overflow-hidden dark:border-dark-3"
                    >
                      {/* Collapsible Header */}
                      <div
                        className="flex items-center justify-between p-4 cursor-pointer bg-gray-50 dark:bg-dark-2 hover:bg-gray-100 dark:hover:bg-dark-3 transition-colors"
                        onClick={() => toggleModuleExpansion(module.id)}
                      >
                        <div className="flex items-center gap-3 flex-1">
                          <svg
                            className={`w-5 h-5 text-dark-6 dark:text-dark-4 transition-transform ${
                              isExpanded ? "rotate-90" : ""
                            }`}
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth={2}
                              d="M9 5l7 7-7 7"
                            />
                          </svg>
                          <div className="flex-1">
                            <h4 className="text-base font-semibold text-dark dark:text-white">
                              Module {moduleIndex + 1}: {module.moduleName}
                            </h4>
                            <p className="text-xs text-dark-6 dark:text-dark-4 mt-0.5">
                              {module.questions.length} question(s)
                              {module.position && ` • ${module.position}`}
                            </p>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            removeImportedModule(module.id);
                          }}
                          className="text-sm text-red hover:text-red/80 px-2 py-1 rounded hover:bg-red/10 transition-colors"
                        >
                          Remove
                        </button>
                      </div>

                      {/* Collapsible Content */}
                      <div
                        className={`overflow-hidden transition-all duration-300 ease-in-out ${
                          isExpanded ? "max-h-[5000px] opacity-100" : "max-h-0 opacity-0"
                        }`}
                      >
                        <div className="p-5 space-y-4">

                          <div className="space-y-2">
                            <label className="block text-body-sm font-medium text-dark dark:text-white">
                              Module Name *
                            </label>
                            <input
                              type="text"
                              value={module.moduleName}
                              onChange={(e) =>
                                updateImportedModule(
                                  module.id,
                                  "moduleName",
                                  e.target.value,
                                )
                              }
                              onClick={(e) => e.stopPropagation()}
                              className="w-full rounded-lg border border-stroke bg-transparent px-4 py-2.5 text-dark outline-none transition focus:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary"
                              required
                            />
                          </div>

                          <div className="space-y-2">
                            <label className="block text-body-sm font-medium text-dark dark:text-white">
                              Position
                            </label>
                            <select
                              value={module.position || ""}
                              onChange={(e) =>
                                updateImportedModule(
                                  module.id,
                                  "position",
                                  e.target.value,
                                )
                              }
                              onClick={(e) => e.stopPropagation()}
                              className="w-full appearance-none rounded-lg border border-stroke bg-transparent px-4 py-2.5 text-dark outline-none transition focus:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary"
                            >
                              <option value="">Select position (optional)</option>
                              <option value="Enterprise Portfolio Manager">
                                Enterprise Portfolio Manager
                              </option>
                              <option value="Project Manager">Project Manager</option>
                              <option value="Program Manager">Program Manager</option>
                              <option value="Portfolio Manager">Portfolio Manager</option>
                              <option value="Business Analyst">Business Analyst</option>
                              <option value="Senior Manager">Senior Manager</option>
                              <option value="Director">Director</option>
                              <option value="Executive">Executive</option>
                              <option value="Team Lead">Team Lead</option>
                              <option value="Consultant">Consultant</option>
                              <option value="Other">Other</option>
                            </select>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-body-sm font-medium text-dark dark:text-white">
                              Description
                            </label>
                            <textarea
                              value={module.moduleDescription}
                              onChange={(e) =>
                                updateImportedModule(
                                  module.id,
                                  "moduleDescription",
                                  e.target.value,
                                )
                              }
                              onClick={(e) => e.stopPropagation()}
                              rows={2}
                              className="w-full rounded-lg border border-stroke bg-transparent px-4 py-2.5 text-dark outline-none transition focus:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary"
                              placeholder="Enter module description"
                            />
                          </div>

                          <div className="space-y-3">
                            <div className="flex items-center justify-between">
                              <label className="block text-body-sm font-medium text-dark dark:text-white">
                                Assessment Questions ({module.questions.length})
                              </label>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  addQuestionToModule(module.id);
                                }}
                                className="text-sm text-primary hover:text-primary/80"
                              >
                                + Add Question
                              </button>
                            </div>
                            {module.questions.map((question, qIndex) => (
                              <div
                                key={question.id}
                                className="flex items-start gap-2 rounded border border-stroke p-3 dark:border-dark-3"
                                onClick={(e) => e.stopPropagation()}
                              >
                                <span className="mt-2 text-sm text-dark-6 dark:text-dark-4">
                                  {qIndex + 1}.
                                </span>
                                <textarea
                                  value={question.question}
                                  onChange={(e) =>
                                    updateImportedModuleQuestion(
                                      module.id,
                                      question.id,
                                      e.target.value,
                                    )
                                  }
                                  rows={2}
                                  className="flex-1 rounded-lg border border-stroke bg-transparent px-3 py-2 text-sm text-dark outline-none transition focus:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary"
                                  placeholder="Enter question text"
                                />
                                {module.questions.length > 1 && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      removeQuestionFromModule(
                                        module.id,
                                        question.id,
                                      )
                                    }
                                    className="mt-2 text-sm text-red hover:text-red/80"
                                  >
                                    Remove
                                  </button>
                                )}
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Module Basic Information - Only show in single module mode */}
          {!isBulkMode && (
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
              />

              <div className="space-y-3">
                <label className="block text-body-sm font-medium text-dark dark:text-white">
                  Position
                </label>
                <select
                  value={position}
                  onChange={(e) => setPosition(e.target.value)}
                  className="w-full appearance-none rounded-lg border border-stroke bg-transparent px-5.5 py-3 text-dark outline-none transition focus:border-primary active:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary [&>option]:text-dark-5 dark:[&>option]:text-dark-6"
                >
                  <option value="">Select position (optional)</option>
                  <option value="Enterprise Portfolio Manager">
                    Enterprise Portfolio Manager
                  </option>
                  <option value="Project Manager">Project Manager</option>
                  <option value="Program Manager">Program Manager</option>
                  <option value="Portfolio Manager">Portfolio Manager</option>
                  <option value="Business Analyst">Business Analyst</option>
                  <option value="Senior Manager">Senior Manager</option>
                  <option value="Director">Director</option>
                  <option value="Executive">Executive</option>
                  <option value="Team Lead">Team Lead</option>
                  <option value="Consultant">Consultant</option>
                  <option value="Other">Other</option>
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
                className="w-full rounded-lg border-[1.5px] border-stroke bg-transparent px-5.5 py-3 text-dark outline-none transition focus:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary"
              />
            </div>
            </div>
          )}

          {/* Assessment Questions Section - Only show in single module mode */}
          {!isBulkMode && (
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
              <button
                type="button"
                onClick={addQuestion}
                className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-6 py-[11px] text-center font-medium text-primary transition hover:bg-primary/10 focus:outline-none"
              >
                Add Assessment Question
              </button>
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
                      className="w-full rounded-lg border-[1.5px] border-stroke bg-transparent px-5.5 py-3 text-dark outline-none transition focus:border-primary dark:border-dark-3 dark:bg-dark-2 dark:text-white dark:focus:border-primary"
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
          )}

          {/* Submit Buttons */}
          <div className="flex justify-end gap-4 pt-4">
            {!isBulkMode && (
              <button
                type="button"
                onClick={() => setShowPreview(!showPreview)}
                className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-10 py-3.5 text-center font-medium text-primary transition hover:bg-primary/10 focus:outline-none lg:px-8 xl:px-10"
              >
                {showPreview ? "Hide Preview" : "Preview"}
              </button>
            )}
            <button
              type="button"
              onClick={() => {
                if (isBulkMode) {
                  setImportedModules([]);
                  setIsBulkMode(false);
                } else {
                  setModuleName("");
                  setModuleDescription("");
                  setPosition("");
                  setQuestions([{ id: "1", question: "" }]);
                }
                setExcelFile(null);
              }}
              className="inline-flex items-center justify-center gap-2.5 rounded-[5px] border border-primary px-10 py-3.5 text-center font-medium text-primary transition hover:bg-primary/10 focus:outline-none lg:px-8 xl:px-10"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={isSubmitting}
              className="inline-flex items-center justify-center gap-2.5 rounded-[5px] bg-[rgb(214,49,41)] px-10 py-3.5 text-center font-medium text-white transition hover:bg-[rgb(214,49,41)]/90 focus:outline-none disabled:opacity-50 disabled:cursor-not-allowed lg:px-8 xl:px-10"
            >
              {isSubmitting
                ? "Submitting..."
                : isBulkMode
                  ? `Create ${importedModules.length} Module(s)`
                  : "Create Module"}
            </button>
          </div>
        </form>

        {/* Preview Section - Only show in single module mode */}
        {!isBulkMode && showPreview && (
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
