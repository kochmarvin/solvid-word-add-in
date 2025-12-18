import * as React from "react";
import { useState } from "react";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string) => void;
}

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
    await props.insertText(text);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  return <div></div>;
};

export default TextInsertion;
