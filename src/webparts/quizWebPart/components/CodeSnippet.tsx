import * as React from 'react';
import { useState, useEffect } from 'react';
import { v4 as uuidv4 } from 'uuid';
import {
  Stack,
  Label,
  TextField,
  Dropdown,
  IDropdownOption,
  DefaultButton,
  PrimaryButton,
  IStackTokens,
  Toggle,
  MessageBar,
  MessageBarType,
  IconButton,
  IIconProps
} from '@fluentui/react';
import { ICodeSnippet, ICodeSnippetProps } from './interfaces';
import styles from './Quiz.module.scss';
import Prism from 'prismjs';
import 'prismjs/components/prism-javascript';
import 'prismjs/components/prism-typescript';
import 'prismjs/components/prism-csharp';
import 'prismjs/components/prism-java';
import 'prismjs/components/prism-python';
import 'prismjs/components/prism-markup';
import 'prismjs/components/prism-css';
import 'prismjs/components/prism-sql';
import 'prismjs/components/prism-bash';
import 'prismjs/components/prism-powershell';
import 'prismjs/plugins/line-numbers/prism-line-numbers';
import 'prismjs/plugins/line-highlight/prism-line-highlight';

// Icons
const removeIcon: IIconProps = { iconName: 'Delete' };
const editIcon: IIconProps = { iconName: 'Edit' };
const saveIcon: IIconProps = { iconName: 'Save' };
const cancelIcon: IIconProps = { iconName: 'Cancel' };

// Stack tokens
const stackTokens: IStackTokens = {
  childrenGap: 10
};

// Language options for the dropdown
const languageOptions: IDropdownOption[] = [
  { key: 'javascript', text: 'JavaScript' },
  { key: 'typescript', text: 'TypeScript' },
  { key: 'csharp', text: 'C#' },
  { key: 'java', text: 'Java' },
  { key: 'python', text: 'Python' },
  { key: 'html', text: 'HTML' },
  { key: 'css', text: 'CSS' },
  { key: 'sql', text: 'SQL' },
  { key: 'bash', text: 'Bash' },
  { key: 'powershell', text: 'PowerShell' }
];

const CodeSnippet: React.FC<ICodeSnippetProps> = (props) => {
  const {
    snippet,
    onChange,
    onRemove,
    isEditing: initialEditMode = false,
    label = "Code Snippet"
  } = props;

  const [code, setCode] = useState<string>(snippet?.code || '');
  const [language, setLanguage] = useState<string>(snippet?.language || 'javascript');
  const [lineNumbers, setLineNumbers] = useState<boolean>(snippet?.lineNumbers || false);
  const [highlightLines, setHighlightLines] = useState<string>(
    snippet?.highlightLines ? snippet.highlightLines.join(',') : ''
  );
  const [isEditing, setIsEditing] = useState<boolean>(initialEditMode || !snippet);
  const [error, setError] = useState<string>('');
  const [currentSnippet, setCurrentSnippet] = useState<ICodeSnippet>(
    snippet || {
      id: uuidv4(),
      code: '',
      language: 'javascript',
      lineNumbers: false
    }
  );

  // Apply syntax highlighting when code or language changes
  useEffect(() => {
    if (!isEditing && code) {
      // Delay to ensure the DOM is ready
      setTimeout(() => {
        Prism.highlightAll();
      }, 0);
    }
  }, [isEditing, code, language]);

  // Update the snippet when form values change
  const updateSnippet = (changes: Partial<ICodeSnippet>): void => {
    const updatedSnippet = {
      ...currentSnippet,
      ...changes
    };
    setCurrentSnippet(updatedSnippet);

    if (!isEditing) {
      onChange(updatedSnippet);
    }
  };

  // Parse highlight lines string into array of numbers
  const parseHighlightLines = (input: string): number[] => {
    if (!input.trim()) return [];

    try {
      // Parse comma-separated values and ranges like "1,3-5,7"
      return input.split(',').flatMap(part => {
        const range = part.trim().split('-');
        if (range.length === 1) {
          const num = parseInt(range[0], 10);
          return isNaN(num) ? [] : [num];
        } else if (range.length === 2) {
          const start = parseInt(range[0], 10);
          const end = parseInt(range[1], 10);
          if (isNaN(start) || isNaN(end)) return [];
          return Array.from({ length: end - start + 1 }, (_, i) => start + i);
        }
        return [];
      });
    } catch (e) {
      console.error('Error parsing highlight lines:', e);
      return [];
    }
  };

  // Save button handler
  const handleSave = (): void => {
    if (!code.trim()) {
      setError('Code snippet cannot be empty');
      return;
    }

    // Validate and parse highlight lines
    const highlightLinesArray = parseHighlightLines(highlightLines);

    const newSnippet: ICodeSnippet = {
      id: currentSnippet.id,
      code: code.trim(),
      language,
      lineNumbers,
      highlightLines: highlightLinesArray
    };

    setCurrentSnippet(newSnippet);
    onChange(newSnippet);
    setIsEditing(false);
    setError('');
  };

  // Start editing
  const handleEdit = (): void => {
    setIsEditing(true);
  };

  // Cancel editing
  const handleCancel = (): void => {
    if (snippet) {
      setCode(snippet.code);
      setLanguage(snippet.language);
      setLineNumbers(snippet.lineNumbers || false);
      setHighlightLines(snippet.highlightLines ? snippet.highlightLines.join(',') : '');
      setCurrentSnippet(snippet);
      setIsEditing(false);
    } else {
      if (onRemove) {
        onRemove();
      }
    }
    setError('');
  };
  

  // Handle code change
  const handleCodeChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    setCode(newValue || '');
    updateSnippet({ code: newValue || '' });
  };

  // Handle language change
  const handleLanguageChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setLanguage(option.key as string);
      updateSnippet({ language: option.key as string });
    }
  };

  // Handle line numbers toggle
  const handleLineNumbersChange = (_: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    setLineNumbers(!!checked);
    updateSnippet({ lineNumbers: !!checked });
  };

  // Handle highlight lines change
  const handleHighlightLinesChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    setHighlightLines(newValue || '');
    updateSnippet({ highlightLines: parseHighlightLines(newValue || '') });
  };

  return (
    <div className={styles.codeSnippetContainer}>
      <Stack tokens={stackTokens}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Label>{label}</Label>

          {!isEditing && (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <IconButton
                iconProps={editIcon}
                title="Edit Snippet"
                ariaLabel="Edit Snippet"
                onClick={handleEdit}
              />
              {onRemove ? (
                <IconButton
                  iconProps={removeIcon}
                  title="Remove Snippet"
                  ariaLabel="Remove Snippet"
                  onClick={onRemove}
                />
              ) : null}

            </Stack>
          )}
        </Stack>

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => setError('')}
            dismissButtonAriaLabel="Close"
          >
            {error}
          </MessageBar>
        )}

        {isEditing ? (
          // Edit mode
          <Stack tokens={stackTokens}>
            <Dropdown
              label="Language"
              selectedKey={language}
              options={languageOptions}
              onChange={handleLanguageChange}
            />

            <TextField
              label="Code"
              multiline
              rows={8}
              value={code}
              onChange={handleCodeChange}
              placeholder="Enter your code here..."
            />

            <Toggle
              label="Line Numbers"
              checked={lineNumbers}
              onChange={handleLineNumbersChange}
            />

            <TextField
              label="Highlight Lines (optional)"
              placeholder="e.g., 1,3-5,7"
              value={highlightLines}
              onChange={handleHighlightLinesChange}
              description="Specify individual lines or ranges (e.g., '1,3-5,7')"
            />

            <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end">
              <DefaultButton
                iconProps={cancelIcon}
                text="Cancel"
                onClick={handleCancel}
              />
              <PrimaryButton
                iconProps={saveIcon}
                text="Save"
                onClick={handleSave}
              />
            </Stack>
          </Stack>
        ) : (
          // Display mode
          <div className={styles.codeDisplay}>
            <pre className={`${lineNumbers ? 'line-numbers' : ''}`}
              data-line={currentSnippet.highlightLines?.join(',')}
              style={{ margin: 0 }}>
              <code className={`language-${language}`}>
                {code}
              </code>
            </pre>
          </div>
        )}
      </Stack>
    </div>
  );
};

export default CodeSnippet;