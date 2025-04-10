import * as React from 'react';
import {
  Stack,
  Text,
  Checkbox,
  TextField,
  ChoiceGroup,
  IChoiceGroupOption,
  IStackTokens
} from '@fluentui/react';
import { IQuizQuestionProps } from './interfaces';
import { QuestionType } from './interfaces';
import styles from './Quiz.module.scss';

// Stack tokens for spacing
const stackTokens: IStackTokens = {
  childrenGap: 15
};

const QuizQuestion: React.FC<IQuizQuestionProps> = (props) => {
  const { question, onAnswerSelect, questionNumber, totalQuestions } = props;

  // Handler for multiple choice and true/false questions
  const handleRadioChange = (
    ev: React.FormEvent<HTMLElement> | undefined,
    option?: IChoiceGroupOption
  ): void => {
    if (option) {
      onAnswerSelect(question.id, option.key);
    }
  };

  // Handler for multiple select questions
  const handleCheckboxChange = (
    ev: React.FormEvent<HTMLElement> | undefined,
    checked?: boolean,
    choiceId?: string
  ): void => {
    if (choiceId === undefined) return;

    const currentSelection = Array.isArray(question.selectedChoice)
      ? [...question.selectedChoice]
      : [];

    let newSelection: string[];
    if (checked) {
      newSelection = [...currentSelection, choiceId];
    } else {
      newSelection = currentSelection.filter(id => id !== choiceId);
    }

    onAnswerSelect(question.id, newSelection);
  };

  // Handler for short answer questions
  const handleTextChange = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement> | undefined,
    newValue?: string
  ): void => {
    onAnswerSelect(question.id, newValue || '');
  };

  // Prepare options for ChoiceGroup
  const getChoiceGroupOptions = (): IChoiceGroupOption[] => {
    return question.choices.map(choice => ({
      key: choice.id,
      text: choice.text
    }));
  };

  // Render different question types
  const renderQuestionContent = (): JSX.Element => {
    switch (question.type) {
      case QuestionType.MultipleChoice:
      case QuestionType.TrueFalse: {
        const options = question.type === QuestionType.TrueFalse
          ? [
            { key: 'true', text: 'True' },
            { key: 'false', text: 'False' }
          ]
          : getChoiceGroupOptions();

        return (
          <ChoiceGroup
            options={options}
            selectedKey={question.selectedChoice as string}
            onChange={handleRadioChange}
            className={styles.choiceGroup}
          />
        );
      }
      case QuestionType.MultiSelect: {
        const selectedChoices = Array.isArray(question.selectedChoice)
          ? question.selectedChoice
          : [];

        return (
          <Stack tokens={stackTokens} className={styles.checkboxGroup}>
            {question.choices.map((choice) => (
              <Checkbox
                key={choice.id}
                label={choice.text}
                checked={selectedChoices.includes(choice.id)}
                onChange={(_, checked) => handleCheckboxChange(_, checked, choice.id)}
              />
            ))}
          </Stack>

        );
      }

      case QuestionType.ShortAnswer:
        return (
          <div className={styles.shortAnswerContainer}>
            <TextField
              value={question.selectedChoice as string || ''}
              onChange={handleTextChange}
              placeholder="Type your answer here"
              multiline={false}
              autoAdjustHeight
              styles={{ fieldGroup: { width: '100%' } }}
            />
            {question.caseSensitive && (
              <Text
                variant="small"
                style={{
                  color: '#6c757d',
                  fontStyle: 'italic',
                  marginTop: '0.5rem'
                }}
              >
                * Answer is case sensitive
              </Text>
            )}
          </div>
        );

      default:
        return (
          <Text>Question type not supported.</Text>
        );
    }
  };

  return (
    <div className={styles.questionCard}>
      <div className={styles.questionHeader}>
        <Text className={styles.questionTitle}>
          {questionNumber}. {question.title}
        </Text>
        <Text className={styles.questionCounter}>
          Question {questionNumber} of {totalQuestions}
        </Text>
      </div>

      <div className={styles.questionInfo}>
        <span className={styles.questionTag}>
          Category: {question.category}
        </span>

        {question.type === QuestionType.ShortAnswer && question.caseSensitive && (
          <span className={styles.questionTag}>
            Case Sensitive
          </span>
        )}

        {question.points && question.points > 1 && (
          <span className={styles.questionTag}>
            {question.points} points
          </span>
        )}
      </div>

      <Stack tokens={stackTokens} className={styles.questionContent}>
        {renderQuestionContent()}
      </Stack>

    </div>
  );
};

export default QuizQuestion;