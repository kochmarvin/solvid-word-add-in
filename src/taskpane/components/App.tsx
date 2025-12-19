/* global Word */
import * as React from "react";
import { useState, useRef } from "react";

import {
  InputGroup,
  InputGroupAddon,
  InputGroupButton,
  InputGroupTextarea,
} from "./ui/input-group";
import { ArrowUpIcon, Brain, SlidersHorizontal, Loader2 } from "lucide-react";
import { Conversation, Message } from "./ui/message";
import { Button } from "./ui/button";
import { generateEditPlan } from "../services/api-service";
import { executeEditPlan } from "../utils/execution-engine";
import { executeSemanticEditPlan } from "../utils/semantic-execution-engine";
import { getSemanticDocument } from "../services/api-service";
import { EditPlan } from "../types/edit-plan";
import { AnchorNotFoundError } from "../utils/errors";

interface ChatMessage {
  role: "user" | "ai";
  content: string;
}

interface PreviewState {
  editPlan: EditPlan;
  response: string;
  semanticEditPlan?: { ops: Array<{ action: string; target_block_id: string; content: string; reason: string }> };
  semanticDocument?: { sections: Array<{ id: string; title: string; level: number; blocks: string[] }>; blocks: Record<string, { type: "paragraph" | "heading"; text: string; level?: number; id?: string }> };
}

interface SelectedRange {
  tag: string;
  text: string;
}

const App: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [preview, setPreview] = useState<PreviewState | null>(null);
  const [selectedRange, setSelectedRange] = useState<SelectedRange | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Check for selected text from context menu on mount
  React.useEffect(() => {
    let cancelled = false;

    const readSettings = () => {
      if (cancelled) return;
      const settings = Office.context?.document?.settings;
      if (!settings) return;

      const selectedText = settings.get("selectedText");
      const selectedTag = settings.get("selectedTag");

      if (selectedText && typeof selectedText === "string" && selectedTag && typeof selectedTag === "string") {
        setInputValue(`Format or edit this text: "${selectedText}"`);
        setSelectedRange({ tag: selectedTag, text: selectedText });

        settings.remove("selectedText");
        settings.remove("selectedTag");
        settings.saveAsync(); // fine with callback omitted

        setTimeout(() => textareaRef.current?.focus(), 100);
      }
    };

    if (typeof Office === "undefined") return () => { };

    Office.onReady(() => {
      readSettings();
    });

    return () => {
      cancelled = true;
    };
  }, []);

  async function readSelectionAndUpdate() {
    if (typeof Word === "undefined") return;

    await Word.run(async (context) => {
      // Find the current solvid-selected Content Control
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();
      
      let foundCC: Word.ContentControl | null = null;
      for (let i = 0; i < contentControls.items.length; i++) {
        const cc = contentControls.items[i];
        cc.load("tag");
        await context.sync();
        
        if (cc.tag && cc.tag.startsWith("solvid-selected-")) {
          foundCC = cc;
          break;
        }
      }
      
      if (foundCC) {
        const ccRange = foundCC.getRange();
        ccRange.load("text");
        await context.sync();
        
        const txt = (ccRange.text || "").trim();
        if (txt) {
          foundCC.load("tag");
          await context.sync();
          setSelectedRange({ tag: foundCC.tag, text: txt });
          setInputValue(`Format or edit this text: "${txt}"`);
        } else {
          setSelectedRange(null);
        }
      } else {
        setSelectedRange(null);
      }
    });
  }

  async function getSelectedText(): Promise<string> {
    if (typeof Word === "undefined") return "";

    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      return (sel.text || "").trim();
    });
  }


  React.useEffect(() => {
    if (typeof Office === "undefined") return () => {}  ;

    let disposed = false;
    let selectionHandlerAttached = false;

    const refreshBadge = async () => {
      if (disposed) return;
      try {
        // Find the current solvid-selected Content Control
        if (typeof Word === "undefined") return;
        
        await Word.run(async (context) => {
          const contentControls = context.document.contentControls;
          contentControls.load("items");
          await context.sync();
          
          let foundCC: Word.ContentControl | null = null;
          for (let i = 0; i < contentControls.items.length; i++) {
            const cc = contentControls.items[i];
            cc.load("tag");
            await context.sync();
            
            if (cc.tag && cc.tag.startsWith("solvid-selected-")) {
              foundCC = cc;
              break;
            }
          }
          
          if (!disposed && foundCC) {
            const ccRange = foundCC.getRange();
            ccRange.load("text");
            foundCC.load("tag");
            await context.sync();
            
            const txt = (ccRange.text || "").trim();
            if (txt) {
              setSelectedRange({ tag: foundCC.tag, text: txt });
            } else {
              setSelectedRange(null);
            }
          } else if (!disposed) {
            setSelectedRange(null);
          }
        });
      } catch {
        if (!disposed) setSelectedRange(null);
      }
    };

    // Debounce so selection drag doesn't spam Word.run
    let t: number | undefined;
    const refreshDebounced = () => {
      if (t) window.clearTimeout(t);
      t = window.setTimeout(() => void refreshBadge(), 150);
    };

    Office.onReady(() => {
      // 1) Refresh once when pane loads
      void refreshBadge();

      // 2) Refresh when user changes selection in the doc
      // Use Office.js DocumentSelectionChanged event (not Word context.document.onSelectionChanged)
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        refreshDebounced,
        () => {
          selectionHandlerAttached = true;
        }
      );
    });

    // 3) Refresh when commands ping via localStorage (right-click menu click)
    const onStorage = (e: StorageEvent) => {
      if (e.key === "solvid:refreshSelection") {
        void refreshBadge();
      }
    };
    window.addEventListener("storage", onStorage);

    return () => {
      disposed = true;
      window.removeEventListener("storage", onStorage);

      if (selectionHandlerAttached) {
        try {
          Office.context.document.removeHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            { handler: refreshDebounced }
          );
        } catch {
          // best-effort cleanup
        }
      }
      if (t) window.clearTimeout(t);
    };
  }, []);


  const handleSubmit = async (e?: React.FormEvent) => {
    e?.preventDefault();

    const prompt = inputValue.trim();
    if (!prompt || isLoading) {
      return;
    }

    // Add user message
    const userMessage: ChatMessage = { role: "user", content: prompt };
    const updatedMessages = [...messages, userMessage];
    setMessages(updatedMessages);
    setInputValue("");
    setIsLoading(true);
    setPreview(null);

    try {
      // Get semantic document structure BEFORE generating the plan
      // This ensures we have the same structure that will be sent to the AI
      const semanticDoc = await getSemanticDocument();
      console.log("handleSubmit - Got semantic document with", Object.keys(semanticDoc.blocks).length, "blocks");
      
      // Call API to generate EditPlan with conversation history and selected range for context
      // Use updatedMessages to include the current user message in conversation history
      const result = await generateEditPlan(prompt, updatedMessages, selectedRange);

      if (!result.ok) {
        // Show error message
        const errorMessage: ChatMessage = {
          role: "ai",
          content: `Error: ${(result as { ok: false; error: { message: string } }).error.message}`,
        };
        setMessages((prev) => [...prev, errorMessage]);
        setIsLoading(false);
        return;
      }

      // Show AI response and preview
      const aiMessage: ChatMessage = {
        role: "ai",
        content: result.response,
      };
      setMessages((prev) => [...prev, aiMessage]);
      
      // Check if this is a semantic edit plan (has ops field)
      const semanticEditPlan = (result as any).semanticEditPlan;
      console.log("handleSubmit - result:", result);
      console.log("handleSubmit - semanticEditPlan:", semanticEditPlan);
      
      if (semanticEditPlan && semanticEditPlan.ops && Array.isArray(semanticEditPlan.ops) && semanticEditPlan.ops.length > 0) {
        console.log("Setting preview with semantic edit plan, ops count:", semanticEditPlan.ops.length);
        console.log("Storing semantic document with", Object.keys(semanticDoc.blocks).length, "blocks");
        setPreview({
          editPlan: result.editPlan,
          response: result.response,
          semanticEditPlan: semanticEditPlan,
          semanticDocument: semanticDoc, // Store the document structure used for generation
        });
      } else {
        console.log("Setting preview with legacy edit plan");
        setPreview({
          editPlan: result.editPlan,
          response: result.response,
        });
      }
    } catch (error) {
      const errorMessage: ChatMessage = {
        role: "ai",
        content: `Unexpected error: ${error instanceof Error ? error.message : String(error)}`,
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleApply = async () => {
    if (!preview) return;

    // Check if Word API is available
    if (typeof Word === "undefined") {
      const errorMsg: ChatMessage = {
        role: "ai",
        content: "Error: Word API is not available. Please ensure the add-in is running in Word.",
      };
      setMessages((prev) => [...prev, errorMsg]);
      return;
    }

    setIsLoading(true);

    try {
      // Check if this is a semantic edit plan (has ops)
      console.log(preview);
      const hasSemanticOps = preview.semanticEditPlan && 
                            preview.semanticEditPlan.ops && 
                            Array.isArray(preview.semanticEditPlan.ops) && 
                            preview.semanticEditPlan.ops.length > 0;
      
      console.log("handleApply - preview.semanticEditPlan:", preview.semanticEditPlan);
      console.log("handleApply - hasSemanticOps:", hasSemanticOps);
      
      if (hasSemanticOps) {
        // Use the semantic document structure from when the plan was generated
        // This ensures block IDs match correctly
        if (!preview.semanticDocument) {
          throw new Error("Semantic document structure not found in preview. This should not happen.");
        }
        
        await executeSemanticEditPlan(
          {
            ops: preview.semanticEditPlan!.ops.map(op => ({
              action: op.action as "insert_after" | "insert_before" | "replace",
              target_block_id: op.target_block_id,
              content: op.content,
              reason: op.reason
            }))
          },
          preview.semanticDocument
        );
        
        const successMessage: ChatMessage = {
          role: "ai",
          content: "Semantic edit plan executed successfully.",
        };
        setMessages((prev) => [...prev, successMessage]);
        setPreview(null);
      } else {
        // Use legacy execution engine
        const result = await executeEditPlan(preview.editPlan);

        if (result.ok) {
          const successMessage: ChatMessage = {
            role: "ai",
            content: result.message,
          };
          setMessages((prev) => [...prev, successMessage]);
          setPreview(null);
        } else {
          const errorResult = result as { ok: false; error_type: string; message: string; details?: Record<string, unknown> };
          let errorMessage = errorResult.message;
          if (errorResult.error_type === "anchor_not_found") {
            const anchor = (errorResult.details as { anchor?: string })?.anchor || "unknown";
            errorMessage = `Anchor not found: ${anchor}. ${errorResult.message}`;
          }
          const errorMsg: ChatMessage = {
            role: "ai",
            content: `Execution failed: ${errorMessage}`,
          };
          setMessages((prev) => [...prev, errorMsg]);
        }
      }
    } catch (error) {
      const errorMsg: ChatMessage = {
        role: "ai",
        content: `Execution error: ${error instanceof Error ? error.message : String(error)}`,
      };
      setMessages((prev) => [...prev, errorMsg]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCancel = () => {
    setPreview(null);
  };

  const renderPreview = () => {
    if (!preview) return null;

    const { editPlan } = preview;
    const actionCount = editPlan.actions.length;
    const blockCounts = editPlan.actions
      .filter((a) => a.type === "replace_section")
      .reduce((sum, a) => sum + (a.type === "replace_section" ? a.blocks.length : 0), 0);

    return (
      <div className="max-w-4xl mx-auto mb-4 p-4 border rounded-lg bg-white dark:bg-gray-800">
        <div className="mb-3">
          <h3 className="font-semibold text-sm mb-2">Edit Plan Preview</h3>
          <div className="text-sm text-gray-600 dark:text-gray-400 space-y-1">
            <div>Actions: {actionCount}</div>
            {blockCounts > 0 && <div>Blocks to insert: {blockCounts}</div>}
            {editPlan.actions.some((a) => a.type === "update_heading_style") && (
              <div>Will update heading styles</div>
            )}
          </div>
        </div>
        <div className="flex gap-2">
          <Button
            onClick={handleApply}
            disabled={isLoading}
            variant="default"
            size="sm"
          >
            {isLoading ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                Applying...
              </>
            ) : (
              "Apply"
            )}
          </Button>
          <Button
            onClick={handleCancel}
            disabled={isLoading}
            variant="outline"
            size="sm"
          >
            Cancel
          </Button>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen p-7 relative pb-32">
      <Conversation>
        {messages.map((msg, index) => (
          <Message key={index} role={msg.role}>
            {msg.content}
          </Message>
        ))}
        {preview && renderPreview()}
      </Conversation>
      <div className="fixed bottom-0 left-0 right-0 w-full p-7 bg-white dark:bg-gray-900">
        <form onSubmit={handleSubmit} className="max-w-4xl mx-auto">
          {selectedRange && (
            <div className="mb-2 flex items-center gap-2">
              <div className="inline-flex items-center gap-1.5 rounded-md bg-blue-100 dark:bg-blue-900/30 px-2.5 py-1 text-xs font-medium text-blue-800 dark:text-blue-200 border border-blue-200 dark:border-blue-800">
                <span className="text-blue-600 dark:text-blue-400">Selected:</span>
                <span className="max-w-[520px] truncate">
                  {selectedRange.text}
                </span>
              </div>

              <button
                type="button"
                onClick={async () => {
                  if (typeof Word !== "undefined") {
                    // Remove ALL active selections by hiding the Content Control borders
                    // Instead of deleting, we'll change appearance to "Hidden" to preserve text
                    await Word.run(async (context) => {
                      const contentControls = context.document.contentControls;
                      contentControls.load("items");
                      await context.sync();
                      
                      // Hide all solvid-selected Content Controls by changing appearance
                      for (let i = 0; i < contentControls.items.length; i++) {
                        const cc = contentControls.items[i];
                        cc.load("tag,appearance");
                        await context.sync();
                        
                        if (cc.tag && cc.tag.startsWith("solvid-selected-")) {
                          // Change appearance to "Hidden" to remove the border
                          // This preserves the text content
                          cc.appearance = "Hidden";
                          // Also change tag to mark as inactive
                          cc.tag = `solvid-inactive-${Date.now()}`;
                        }
                      }
                      await context.sync();
                    });
                  }
                  setSelectedRange(null);
                }}
                className="text-xs text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 underline"
              >
                Clear
              </button>

              <button
                type="button"
                onClick={async () => {
                  await readSelectionAndUpdate();
                }}
                className="text-xs text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 underline"
              >
                Update
              </button>
            </div>
          )}
          <InputGroup>
          <InputGroupTextarea
              ref={textareaRef}
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  handleSubmit();
                }
              }}
              placeholder="Describe the document changes you want..."
              disabled={isLoading}
          />
          <InputGroupAddon align="block-end">
              <InputGroupButton variant="outline" size="sm" type="button">
              <Brain />
              Wissen verwalten
            </InputGroupButton>
              <InputGroupButton variant="outline" size="sm" type="button">
              <SlidersHorizontal />
            </InputGroupButton>
            <InputGroupButton
              variant="default"
              className="rounded-full ml-auto"
              size="icon-sm"
                type="submit"
                disabled={isLoading || !inputValue.trim()}
            >
                {isLoading ? (
                  <Loader2 className="h-4 w-4 animate-spin" />
                ) : (
              <ArrowUpIcon />
                )}
              <span className="sr-only">Send</span>
            </InputGroupButton>
          </InputGroupAddon>
        </InputGroup>
        </form>
      </div>
    </div>
  );
};

export default App;
