import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Stack,
  Text,
  Checkbox,
  TextField,
  ChoiceGroup,
  IChoiceGroupOption,
  IStackTokens,
  Image,
  ImageFit,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { IQuizQuestionProps } from './interfaces';
import { QuestionType } from './interfaces';
import styles from './Quiz.module.scss';
import QuestionTimer from './QuestionTimer';

// Import Prism for code highlighting
import Prism from 'prismjs';
import { MatchingQuestion } from './MatchingQuestion';

// Stack tokens for spacing
const stackTokens: IStackTokens = {
  childrenGap: 15
};

const QuizQuestion: React.FC<IQuizQuestionProps> = (props) => {
  const { question, onAnswerSelect, questionNumber, totalQuestions } = props;
  const [timeLimitExpired, setTimeLimitExpired] = useState(false);

  // Apply syntax highlighting to code snippets when component mounts or changes
  useEffect(() => {
    if (question.codeSnippets && question.codeSnippets.length > 0) {
      setTimeout(() => {
        Prism.highlightAll();
      }, 100);
    }
  }, [question.codeSnippets]);

  // Handler for multiple choice and true/false questions
  const handleRadioChange = (
    _ev: React.FormEvent<HTMLElement> | undefined,
    option?: IChoiceGroupOption
  ): void => {
    if (option) {
      onAnswerSelect(question.id, option.key);
    }
  };

  // Handler for multiple select questions
  const handleCheckboxChange = (
    _ev: React.FormEvent<HTMLElement> | undefined,
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
    _ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement> | undefined,
    newValue?: string
  ): void => {
    onAnswerSelect(question.id, newValue || '');
  };

  // Handle time expired
  const handleTimeExpired = (): void => {
    setTimeLimitExpired(true);
    // If the parent provided a time expired handler, call it
    if (props.onTimeExpired) {
      props.onTimeExpired(question.id);
    }
  };

  // Prepare options for ChoiceGroup
  const getChoiceGroupOptions = (): IChoiceGroupOption[] => {
    return question.choices.map(choice => {
      // Check if choice has an image
      if (choice.image) {
        return {
          key: choice.id,
          text: choice.text,
          imageSrc: choice.image.url,
          imageAlt: choice.image.altText || choice.text
        };
      }
      return {
        key: choice.id,
        text: choice.text
      };
    });
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
              <div key={choice.id} className={styles.multiSelectChoice}>
                <Checkbox
                  label={choice.text}
                  checked={selectedChoices.includes(choice.id)}
                  onChange={(ev, checked) => handleCheckboxChange(ev, checked, choice.id)}
                />
                {choice.image && (
                  <div className={styles.choiceImage}>
                    <Image
                      src={choice.image.url}
                      alt={choice.image.altText || choice.text}
                      width={200}
                      height={120}
                      imageFit={ImageFit.contain}
                    />
                  </div>
                )}
              </div>
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
        case QuestionType.Matching:
          return (
            <div className={styles.matchingContainer}>
              <MatchingQuestion
                questionId={question.id}
                matchingPairs={question.matchingPairs || []}
                onMatchingPairsChange={(questionId, updatedPairs) => {
                  // Convert matching pairs to the selected choices format
                  // This helps maintain compatibility with the existing answer handling logic
                  const selectedPairIds = updatedPairs
                    .map(pair => `${pair.id}:${pair.userSelectedRightId || ''}`)
                    .filter(mapping => mapping.endsWith(':') === false);
                  
                  onAnswerSelect(questionId, selectedPairIds);
                }}
                disabled={timeLimitExpired}
              />
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

      {/* Display question timer if timeLimit is set */}
      {question.timeLimit && question.timeLimit > 0 && (
        <div className={styles.questionTimerWrapper}>
          <QuestionTimer 
            timeLimit={question.timeLimit} 
            onTimeExpired={handleTimeExpired}
            paused={false}
          />
        </div>
      )}

      {/* Display time expired message */}
      {timeLimitExpired && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline={false}
          className={styles.timeExpiredMessage}
        >
          Time limit reached for this question. Your answer may not be counted.
        </MessageBar>
      )}

      {question.description && (
        <div className={styles.questionDescription}>
          <RichText 
            value={question.description}
            isEditMode={false}
          />
        </div>
      )}

      {/* Display question images */}
      {question.images && question.images.length > 0 && (
        <div className={styles.questionImages}>
          {question.images.map(image => (
            <div key={image.id} className={styles.questionImage}>
              <Image
                src={image.url}
                alt={image.altText || 'Question image'}
                width={image.width || 500}
                height={image.height || 300}
                imageFit={ImageFit.contain}
              />
              {image.altText && (
                <Text className={styles.imageCaption}>{image.altText}</Text>
              )}
            </div>
          ))}
        </div>
      )}

      {/* Display code snippets */}
      {question.codeSnippets && question.codeSnippets.length > 0 && (
        <div className={styles.questionCodeSnippets}>
          {question.codeSnippets.map(snippet => (
            <div key={snippet.id} className={styles.codeDisplay}>
              <pre className={snippet.lineNumbers ? 'line-numbers' : ''} 
                  data-line={snippet.highlightLines?.join(',')}>
                <code className={`language-${snippet.language}`}>
                  {snippet.code}
                </code>
              </pre>
            </div>
          ))}
        </div>
      )}

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