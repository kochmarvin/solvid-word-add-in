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
import { EditPlan } from "../types/edit-plan";
import { AnchorNotFoundError } from "../utils/errors";

interface ChatMessage {
  role: "user" | "ai";
  content: string;
}

interface PreviewState {
  editPlan: EditPlan;
  response: string;
}

const App: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [preview, setPreview] = useState<PreviewState | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Check for selected text from context menu on mount
  React.useEffect(() => {
    if (typeof Office !== "undefined" && Office.context?.document?.settings) {
      const selectedText = Office.context.document.settings.get("selectedText");
      if (selectedText && typeof selectedText === "string") {
        // Pre-fill the input with a prompt about the selected text
        setInputValue(`Format or edit this text: "${selectedText}"`);
        // Clear the setting so it doesn't persist
        Office.context.document.settings.remove("selectedText");
        Office.context.document.settings.saveAsync();
        // Focus the textarea
        setTimeout(() => {
          textareaRef.current?.focus();
        }, 100);
      }
    }
  }, []);

  const handleSubmit = async (e?: React.FormEvent) => {
    e?.preventDefault();

    const prompt = inputValue.trim();
    if (!prompt || isLoading) {
      return;
    }

    // Add user message
    const userMessage: ChatMessage = { role: "user", content: prompt };
    setMessages((prev) => [...prev, userMessage]);
    setInputValue("");
    setIsLoading(true);
    setPreview(null);

    try {
      // Call API to generate EditPlan with conversation history for context
      const result = await generateEditPlan(prompt, messages);

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
      setPreview({
        editPlan: result.editPlan,
        response: result.response,
      });
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
