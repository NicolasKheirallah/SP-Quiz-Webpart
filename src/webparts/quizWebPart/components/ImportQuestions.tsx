import * as React from 'react';
import { useState } from 'react';
import { IImportQuestionsProps, IQuizQuestion, IChoice, QuestionType } from './interfaces';
import { v4 as uuidv4 } from 'uuid';
import {
  Button,
  Input,
  Text,
  Textarea,
  MessageBar,
  MessageBarBody,
  Dropdown,
  Option,
  Field,
  Label,
  Card,
  Divider
} from '@fluentui/react-components';
import { 
  ArrowDownloadRegular, 
  DocumentRegular, 
  InfoRegular
} from '@fluentui/react-icons';
import styles from './Quiz.module.scss';

const ImportQuestions: React.FC<IImportQuestionsProps> = (props) => {
  const { onImportQuestions, onCancel, existingCategories } = props;
  
  const [csvContent, setCsvContent] = useState<string>('');
  const [jsonContent, setJsonContent] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [successMessage, setSuccessMessage] = useState<string>('');
  const [importFormat, setImportFormat] = useState<string>('csv');
  const [defaultCategory, setDefaultCategory] = useState<string>('');
  const [newCategory, setNewCategory] = useState<string>('');
  
  // Sample template content
  const csvTemplate = `"Question","Category","Type","Option 1","Option 2","Option 3","Option 4","Correct Answer"
"What is SharePoint?","SharePoint Basics","multipleChoice","A document management system","A social network","A database system","An operating system","1"
"SharePoint is developed by which company?","SharePoint Basics","multipleChoice","Google","Microsoft","Oracle","IBM","2"`;

  const jsonTemplate = `[
  {
    "title": "What is SharePoint?",
    "category": "SharePoint Basics",
    "type": "multipleChoice",
    "choices": [
      { "id": "${uuidv4()}", "text": "A document management system", "isCorrect": true },
      { "id": "${uuidv4()}", "text": "A social network", "isCorrect": false },
      { "id": "${uuidv4()}", "text": "A database system", "isCorrect": false },
      { "id": "${uuidv4()}", "text": "An operating system", "isCorrect": false }
    ]
  },
  {
    "title": "SharePoint is developed by which company?",
    "category": "SharePoint Basics", 
    "type": "multipleChoice",
    "choices": [
      { "id": "${uuidv4()}", "text": "Google", "isCorrect": false },
      { "id": "${uuidv4()}", "text": "Microsoft", "isCorrect": true },
      { "id": "${uuidv4()}", "text": "Oracle", "isCorrect": false },
      { "id": "${uuidv4()}", "text": "IBM", "isCorrect": false }
    ]
  }
]`;

  // Handle file upload
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const content = e.target?.result as string;
      
      if (file.name.endsWith('.csv')) {
        setCsvContent(content);
        setImportFormat('csv');
      } else if (file.name.endsWith('.json')) {
        setJsonContent(content);
        setImportFormat('json');
      } else {
        setErrorMessage('Unsupported file format. Please upload a CSV or JSON file.');
      }
    };
    
    reader.readAsText(file);
  };
  
  // Parse CSV content
  const parseCSV = (csv: string): IQuizQuestion[] => {
    try {
      // Simple CSV parser - in a real app, use a library like PapaParse
      const lines = csv.split(/\\r?\\n/);
      
      // Extract headers
      const headers = lines[0].split(',').map(h => 
        h.trim().replace(/^"|"$/g, '')
      );
      
      const questionIndex = headers.findIndex(h => h.toLowerCase() === 'question');
      const categoryIndex = headers.findIndex(h => h.toLowerCase() === 'category');
      const typeIndex = headers.findIndex(h => h.toLowerCase() === 'type');
      const correctAnswerIndex = headers.findIndex(h => h.toLowerCase() === 'correct answer');
      
      // Find option columns
      const optionIndices: number[] = [];
      for (let i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase().startsWith('option')) {
          optionIndices.push(i);
        }
      }
      
      // Parse data rows
      const questions: IQuizQuestion[] = [];
      for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue; // Skip empty lines
        
        // Split by comma but respect quotes
        const values: string[] = [];
        let inQuotes = false;
        let currentValue = '';
        
        for (let j = 0; j < lines[i].length; j++) {
          const char = lines[i][j];
          
          if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === ',' && !inQuotes) {
            values.push(currentValue.replace(/^"|"$/g, ''));
            currentValue = '';
          } else {
            currentValue += char;
          }
        }
        
        // Add the last value
        values.push(currentValue.replace(/^"|"$/g, ''));
        
        // Create question object
        const questionTitle = questionIndex >= 0 ? values[questionIndex] : '';
        let questionCategory = categoryIndex >= 0 ? values[categoryIndex] : '';
        
        // Use default category if specified and the row doesn't have one
        if (!questionCategory && defaultCategory) {
          questionCategory = defaultCategory;
        }
        
        // Determine question type
        let questionType = QuestionType.MultipleChoice; // Default
        if (typeIndex >= 0) {
          const typeValue = values[typeIndex].toLowerCase();
          
          if (typeValue === 'truefaise' || typeValue === 'true/false') {
            questionType = QuestionType.TrueFalse;
          } else if (typeValue === 'multiselect') {
            questionType = QuestionType.MultiSelect;
          } else if (typeValue === 'shortanswer') {
            questionType = QuestionType.ShortAnswer;
          }
        }
        
        // Parse options and determine correct answers
        const choices: IChoice[] = [];
        const correctAnswer = correctAnswerIndex >= 0 ? values[correctAnswerIndex] : '';
        
        for (let j = 0; j < optionIndices.length; j++) {
          const optionIndex = optionIndices[j];
          const optionText = values[optionIndex];
          
          if (optionText) {
            let isCorrect = false;
            
            if (correctAnswer) {
              // Handle different correct answer formats
              if (correctAnswer === optionText) {
                isCorrect = true;
              } else if (correctAnswer === (j+1).toString()) {
                isCorrect = true;
              } else if (correctAnswer.split(',').includes((j+1).toString())) {
                isCorrect = true;
              }
            }
            
            choices.push({
              id: uuidv4(),
              text: optionText,
              isCorrect
            });
          }
        }
        
        // Create the question
        questions.push({
          id: Date.now() + i,
          title: questionTitle,
          category: questionCategory,
          type: questionType,
          choices: choices,
          correctAnswer: questionType === QuestionType.ShortAnswer ? correctAnswer : undefined
        });
      }
      
      return questions;
    } catch (error) {
      console.error('CSV parsing error:', error);
      setErrorMessage('Error parsing CSV. Please check the format and try again.');
      return [];
    }
  };
  
  // Parse JSON content
  const parseJSON = (json: string): IQuizQuestion[] => {
    try {
      const parsedQuestions = JSON.parse(json) as IQuizQuestion[];
      
      // Validate and sanitize
      return parsedQuestions.map(q => {
        // Ensure each question has an id and required fields
        let category = q.category || '';
        
        // Use default category if specified and the question doesn't have one
        if (!category && defaultCategory) {
          category = defaultCategory;
        }
        
        // Ensure choices have IDs
        const choices = (q.choices || []).map(c => ({
          id: c.id || uuidv4(),
          text: c.text || '',
          isCorrect: !!c.isCorrect
        }));
        
        return {
          id: q.id || Date.now() + Math.floor(Math.random() * 1000),
          title: q.title || '',
          category,
          type: q.type || QuestionType.MultipleChoice,
          choices,
          correctAnswer: q.correctAnswer,
          explanation: q.explanation
        };
      });
    } catch (error) {
      console.error('JSON parsing error:', error);
      setErrorMessage('Error parsing JSON. Please check the format and try again.');
      return [];
    }
  };
  
  // Handle import submission
  const handleImport = (): void => {
    try {
      setErrorMessage('');
      
      let importedQuestions: IQuizQuestion[] = [];
      
      if (importFormat === 'csv' && csvContent) {
        importedQuestions = parseCSV(csvContent);
      } else if (importFormat === 'json' && jsonContent) {
        importedQuestions = parseJSON(jsonContent);
      } else {
        setErrorMessage('Please provide content to import.');
        return;
      }
      
      if (importedQuestions.length === 0) {
        setErrorMessage('No valid questions found to import.');
        return;
      }
      
      setSuccessMessage(`Successfully parsed ${importedQuestions.length} questions.`);
      onImportQuestions(importedQuestions);
      
    } catch (error) {
      console.error('Import error:', error);
      setErrorMessage('Error importing questions. Please try again.');
    }
  };
  
  // Handle using template
  const useTemplate = (format: 'csv' | 'json'): void => {
    if (format === 'csv') {
      setCsvContent(csvTemplate);
      setImportFormat('csv');
    } else {
      setJsonContent(jsonTemplate);
      setImportFormat('json');
    }
  };

  return (
    <div className={styles.importQuestionsForm}>
      <Card>
        <Text weight="semibold" size={500}>Import Questions</Text>
        <Text>Import questions from CSV or JSON format</Text>
        
        {errorMessage && (
          <MessageBar intent="error">
            <MessageBarBody>{errorMessage}</MessageBarBody>
          </MessageBar>
        )}
        
        {successMessage && (
          <MessageBar intent="success">
            <MessageBarBody>{successMessage}</MessageBarBody>
          </MessageBar>
        )}
        
        <div className={styles.importOptions}>
          <Field label="Import Format">
            <Dropdown
              value={importFormat}
              onOptionSelect={(_, data) => setImportFormat(data.optionValue || 'csv')}
            >
              <Option value="csv">CSV Format</Option>
              <Option value="json">JSON Format</Option>
            </Dropdown>
          </Field>
          
          <Field label="Default Category (optional)">
            <Dropdown
              value={defaultCategory}
              onOptionSelect={(_, data) => {
                const value = data.optionValue || '';
                setDefaultCategory(value);
                if (value === 'new') {
                  setNewCategory('');
                }
              }}
            >
              <Option value="">None</Option>
              {existingCategories.map(cat => (
                <Option key={cat} value={cat}>{cat}</Option>
              ))}
              <Option value="new">Add new category</Option>
            </Dropdown>
          </Field>
          
          {defaultCategory === 'new' && (
            <Field label="New Category Name">
              <Input
                value={newCategory}
                onChange={(_, data) => setNewCategory(data.value)}
                placeholder="Enter new category name"
              />
            </Field>
          )}
        </div>
        
        <div className={styles.importFileSection}>
          <Label>Upload File</Label>
          <input
            type="file"
            accept={importFormat === 'csv' ? '.csv' : '.json'}
            onChange={handleFileUpload}
            className={styles.nativeFileInput}
          />
          <Text size={200}>or</Text>
          <Button
            appearance="subtle"
            onClick={() => useTemplate(importFormat as 'csv' | 'json')}
            icon={<DocumentRegular />}
          >
            Use {importFormat.toUpperCase()} Template
          </Button>
        </div>
        
        <Divider />
        
        {importFormat === 'csv' ? (
          <Field label="CSV Content">
            <Textarea
              value={csvContent}
              onChange={(_, data) => setCsvContent(data.value)}
              placeholder="Paste CSV content here..."
              resize="vertical"
              style={{ minHeight: '200px' }}
            />
          </Field>
        ) : (
          <Field label="JSON Content">
            <Textarea
              value={jsonContent}
              onChange={(_, data) => setJsonContent(data.value)}
              placeholder="Paste JSON content here..."
              resize="vertical"
              style={{ minHeight: '200px' }}
            />
          </Field>
        )}
        
        <MessageBar intent="info">
          <MessageBarBody>
            <InfoRegular />
            {importFormat === 'csv' 
              ? 'CSV format should have columns for Question, Category, Type, Options, and Correct Answer.' 
              : 'JSON should be an array of question objects with title, category, type, and choices properties.'}
          </MessageBarBody>
        </MessageBar>
        
        <div className={styles.formButtons}>
          <Button
            appearance="primary"
            onClick={handleImport}
            icon={<ArrowDownloadRegular />}
          >
            Import Questions
          </Button>
          <Button
            appearance="secondary"
            onClick={onCancel}
          >
            Cancel
          </Button>
        </div>
      </Card>
    </div>
  );
};

export default ImportQuestions;