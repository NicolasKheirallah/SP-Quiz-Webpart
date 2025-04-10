import * as React from 'react';
import { useState } from 'react';
import { IImportQuestionsProps, IQuizQuestion, IChoice, QuestionType } from './interfaces';
import { v4 as uuidv4 } from 'uuid';
import {
    Dialog,
    DialogType,
    DialogFooter,
    PrimaryButton,
    DefaultButton,
    TextField,
    Label,
    MessageBar,
    MessageBarType,
    Dropdown,
    IDropdownOption,
    Stack,
    IStackTokens,
    IIconProps,
    Text,
    Spinner,
    SpinnerSize,
    Link,
    ChoiceGroup,
    IChoiceGroupOption,
    Icon
} from '@fluentui/react';
import styles from './Quiz.module.scss';

// Icons
const uploadIcon: IIconProps = { iconName: 'Upload' };
const fileIcon: IIconProps = { iconName: 'DocumentSet' };
const templateIcon: IIconProps = { iconName: 'FileTemplate' };

// Stack tokens for spacing
const stackTokens: IStackTokens = {
    childrenGap: 15
};

const ImportQuestionsDialog: React.FC<IImportQuestionsProps> = (props) => {
    const { onImportQuestions, onCancel, existingCategories } = props;

    const [csvContent, setCsvContent] = useState<string>('');
    const [jsonContent, setJsonContent] = useState<string>('');
    const [errorMessage, setErrorMessage] = useState<string>('');
    const [successMessage, setSuccessMessage] = useState<string>('');
    const [importFormat, setImportFormat] = useState<string>('json');
    const [defaultCategory, setDefaultCategory] = useState<string>('');
    const [newCategory, setNewCategory] = useState<string>('');
    const [isProcessing, setIsProcessing] = useState<boolean>(false);
    const [uploadedFile, setUploadedFile] = useState<File | null>(null);

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

    // Format options for the ChoiceGroup
    const formatOptions: IChoiceGroupOption[] = [
        { key: 'json', text: 'JSON Format', iconProps: { iconName: 'FileCode' } },
        { key: 'csv', text: 'CSV Format', iconProps: { iconName: 'ExcelDocument' } }
    ];

    // Handle file upload
    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>): void => {
        const file = event.target.files?.[0];
        if (!file) return;

        setUploadedFile(file);

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
            setIsProcessing(true);
            // Simple CSV parser - in a real app, you might use a library like PapaParse
            const lines = csv.split(/\r?\n/);

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
                    questionCategory = defaultCategory === 'new' ? newCategory : defaultCategory;
                }

                // Determine question type
                let questionType = QuestionType.MultipleChoice; // Default
                if (typeIndex >= 0) {
                    const typeValue = values[typeIndex].toLowerCase();

                    if (typeValue === 'truefalse' || typeValue === 'true/false') {
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
                            } else if (correctAnswer === (j + 1).toString()) {
                                isCorrect = true;
                            } else if (correctAnswer.split(',').includes((j + 1).toString())) {
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
                    correctAnswer: questionType === QuestionType.ShortAnswer ? correctAnswer : undefined,
                    lastModified: new Date().toISOString()
                });
            }

            return questions;
        } catch (error) {
            console.error('CSV parsing error:', error);
            setErrorMessage('Error parsing CSV. Please check the format and try again.');
            return [];
        } finally {
            setIsProcessing(false);
        }
    };

    // Parse JSON content
    const parseJSON = (json: string): IQuizQuestion[] => {
        try {
            setIsProcessing(true);
            const parsedQuestions = JSON.parse(json) as IQuizQuestion[];

            // Validate and sanitize
            return parsedQuestions.map(q => {
                // Ensure each question has an id and required fields
                let category = q.category || '';

                // Use default category if specified and the question doesn't have one
                if (!category && defaultCategory) {
                    category = defaultCategory === 'new' ? newCategory : defaultCategory;
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
                    explanation: q.explanation,
                    lastModified: new Date().toISOString()
                };
            });
        } catch (error) {
            console.error('JSON parsing error:', error);
            setErrorMessage('Error parsing JSON. Please check the format and try again.');
            return [];
        } finally {
            setIsProcessing(false);
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
    const useTemplate = (): void => {
        if (importFormat === 'csv') {
            setCsvContent(csvTemplate);
        } else {
            setJsonContent(jsonTemplate);
        }
    };

    // Generate category dropdown options
    const categoryOptions: IDropdownOption[] = [
        { key: '', text: 'None' },
        ...existingCategories.map(cat => ({ key: cat, text: cat })),
        { key: 'new', text: 'Add new category' }
    ];

    return (
        <Dialog
            hidden={false}
            onDismiss={onCancel}
            dialogContentProps={{
                type: DialogType.largeHeader,
                title: 'Import Questions'
            }}
            modalProps={{
                isBlocking: true,
                styles: { main: { maxWidth: '700px', minWidth: '600px' } }
            }}
        >
            {errorMessage && (
                <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    dismissButtonAriaLabel="Close"
                    styles={{ root: { marginBottom: 15 } }}
                    className={`${styles.statusBar} ${styles.error}`}
                >
                    {errorMessage}
                </MessageBar>
            )}

            {successMessage && (
                <MessageBar
                    messageBarType={MessageBarType.success}
                    isMultiline={false}
                    dismissButtonAriaLabel="Close"
                    styles={{ root: { marginBottom: 15 } }}
                    className={`${styles.statusBar} ${styles.success}`}
                >
                    {successMessage}
                </MessageBar>
            )}

            <Stack tokens={stackTokens} className={styles.formWrapper}>
                <ChoiceGroup
                    options={formatOptions}
                    selectedKey={importFormat}
                    onChange={(_, option) => option && setImportFormat(option.key)}
                    label="Import Format"
                />

                <Dropdown
                    label="Default Category (Optional)"
                    selectedKey={defaultCategory}
                    onChange={(_, option) => {
                        if (option) {
                            setDefaultCategory(option.key as string);
                            if (option.key === 'new') {
                                setNewCategory('');
                            }
                        }
                    }}
                    options={categoryOptions}
                    placeholder="Select a default category"
                />

                {defaultCategory === 'new' && (
                    <TextField
                        label="New Category Name"
                        required
                        value={newCategory}
                        onChange={(_, value) => setNewCategory(value || '')}
                        placeholder="Enter new category name"
                    />
                )}

                <div className={styles.importFileSection}>
                    <Label>
                    <Icon {...fileIcon} style={{ marginRight: '8px' }} />
                    Upload {importFormat.toUpperCase()} File
                    </Label>
                    <Stack horizontal verticalAlign="center">
                        <input
                            type="file"
                            accept={importFormat === 'csv' ? '.csv' : '.json'}
                            onChange={handleFileUpload}
                            style={{ marginBottom: '10px' }}
                            className={styles.nativeFileInput}
                        />
                    </Stack>
                    <Text>{uploadedFile ? `Selected file: ${uploadedFile.name}` : 'No file selected'}</Text>
                    <Text style={{ margin: '10px 0' }}>- OR -</Text>
                    <Link onClick={useTemplate}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                        <Icon {...templateIcon} />
                        <Text>Use {importFormat.toUpperCase()} Template</Text>
                        </Stack>
                    </Link>
                </div>

                <Stack className={`${styles.statusBar} ${styles.info}`}>
                    <Text style={{ fontWeight: 600 }}>{importFormat === 'csv' ? 'CSV Format Guidelines' : 'JSON Format Guidelines'}</Text>
                    {importFormat === 'csv' ? (
                        <Text>
                            CSV files should include headers for &quot;Question&quot;, &quot;Category&quot;, &quot;Type&quot;, &quot;Option 1&quot;, &quot;Option 2&quot;, etc.,
                            and &quot;Correct Answer&quot;. The correct answer can be the option number or exact option text.
                        </Text>
                    ) : (
                        <Text>
                            JSON should be an array of question objects with properties like &quot;title&quot;, &quot;category&quot;, &quot;type&quot;,
                            and &quot;choices&quot; (an array of choice objects with &quot;text&quot; and &quot;isCorrect&quot; properties).
                        </Text>
                    )}
                </Stack>

                <TextField
                    label={importFormat === 'csv' ? 'CSV Content' : 'JSON Content'}
                    multiline
                    rows={8}
                    value={importFormat === 'csv' ? csvContent : jsonContent}
                    onChange={(_, value) => importFormat === 'csv' ? setCsvContent(value || '') : setJsonContent(value || '')}
                    placeholder={`Paste ${importFormat.toUpperCase()} content here...`}
                />
            </Stack>

            <DialogFooter className={styles.formButtons}>
                {isProcessing ? (
                    <Spinner label="Processing..." size={SpinnerSize.medium} />
                ) : (
                    <>
                        <PrimaryButton
                            onClick={handleImport}
                            text="Import Questions"
                            disabled={importFormat === 'csv' ? !csvContent : !jsonContent}
                            iconProps={uploadIcon}
                        />
                        <DefaultButton onClick={onCancel} text="Cancel" />
                    </>
                )}
            </DialogFooter>
        </Dialog>
    );
};

export default ImportQuestionsDialog;