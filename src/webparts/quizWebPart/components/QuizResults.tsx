import * as React from 'react';
import { IQuizResultsProps } from './interfaces';
import {
  Stack,
  IStackTokens,
  Text,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  ITextStyles,
  FontWeights,
  IIconProps,
  Pivot,
  PivotItem,
  Separator,
  Label
} from '@fluentui/react';
import { Accordion } from "@pnp/spfx-controls-react/lib/Accordion";
import { ListView, IViewField, SelectionMode } from "@pnp/spfx-controls-react/lib/ListView";
import styles from './Quiz.module.scss';
const retakeIcon: IIconProps = { iconName: 'Refresh' };
const reviewIcon: IIconProps = { iconName: 'ReadingMode' };
const excellentIconName = 'Trophy';
const goodIconName = 'Emoji2';
const averageIconName = 'EmojiNeutral';
const poorIconName = 'Sad';
const checkIconName = 'CheckMark';

// Styles
const resultsTitleStyles: ITextStyles = {
  root: {
    fontSize: '24px',
    fontWeight: FontWeights.semibold,
    marginBottom: '20px'
  }
};

const scoreValueStyles: ITextStyles = {
  root: {
    fontSize: '42px',
    fontWeight: FontWeights.bold,
    color: 'white'
  }
};

const scoreTextStyles: ITextStyles = {
  root: {
    fontSize: '16px',
    color: 'white',
    marginTop: '4px'
  }
};

const scoreDetailsStyles: ITextStyles = {
  root: {
    fontSize: '18px',
    marginBottom: '16px'
  }
};

const resultMessageStyles: ITextStyles = {
  root: {
    fontSize: '16px',
    fontStyle: 'italic',
    marginBottom: '16px'
  }
};

// Get appropriate icon based on score
const getScoreIconName = (percentage: number): string => {
  if (percentage >= 90) return excellentIconName;
  if (percentage >= 70) return goodIconName;
  if (percentage >= 50) return averageIconName;
  return poorIconName;
};

const QuizResults: React.FC<IQuizResultsProps> = (props) => {
  const {
    score,
    totalQuestions,
    totalPoints,
    isSubmitting,
    submissionSuccess,
    submissionError,
    onRetakeQuiz,
    messages,
    detailedResults // Ensure this contains { score, totalPoints, percentage, questionResults: [...] }
  } = props;

  const [activeView, setActiveView] = React.useState<string>('summary');

  // Calculate percentage based on props directly if detailedResults isn't always present initially
  const summaryPercentage = totalQuestions > 0 ? Math.round((score / totalPoints) * 100) : 0;

  // Determine result message based on summary percentage
  let resultMessage = '';
  if (summaryPercentage >= 90) {
    resultMessage = messages.excellent || 'Excellent! You have mastered this topic!';
  } else if (summaryPercentage >= 70) {
    resultMessage = messages.good || 'Good job! You have a solid understanding.';
  } else if (summaryPercentage >= 50) {
    resultMessage = messages.average || 'Not bad. There\'s room for improvement.';
  } else {
    resultMessage = messages.poor || 'Keep studying. You\'ll get better with practice.';
  }

  // Stack tokens for spacing
  const stackTokens: IStackTokens = {
    childrenGap: 16
  };

  // Get performance message based on percentage - THIS IS USED in detailed view
  const getPerformanceMessage = (percentage: number): string => {
    if (percentage >= 90) return 'Excellent! You\'ve mastered this material.';
    if (percentage >= 75) return 'Great job! You have a good understanding of the material.';
    if (percentage >= 60) return 'Good effort! With a bit more study, you\'ll master this.';
    if (percentage >= 40) return 'Keep practicing. You\'re making progress.';
    return 'Don\'t give up! With more practice, you\'ll improve.';
  };

  // Format answers for display
  const formatAnswer = (answer: string | string[] | undefined): string => {
    if (answer === undefined || answer === null) return 'No answer provided'; // Handle null too

    if (Array.isArray(answer)) {
      return answer.length > 0 ? answer.join(', ') : 'No answer provided';
    }
    return answer.toString().trim() !== '' ? answer.toString() : 'No answer provided';
  };



  // Configure ListView columns for detailed results
  const viewFields: IViewField[] = React.useMemo(() => [
    {
      name: 'title',
      displayName: 'Question',
      minWidth: 200,
      maxWidth: 300,
      // Ensure item.title exists and is a string
      render: (item) => <Text>{item?.title || 'N/A'}</Text>
    },
    {
      name: 'isCorrect',
      displayName: 'Result',
      minWidth: 80, // Adjusted width
      maxWidth: 100,
      render: (item) => (
        item && typeof item.isCorrect === 'boolean' ? ( // Add check for item and isCorrect
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
            <Icon
              iconName={item.isCorrect ? 'CheckMark' : 'Cancel'}
              style={{
                color: item.isCorrect ? '#107c10' : '#d13438',
              }}
            />
            <Text>{item.isCorrect ? 'Correct' : 'Incorrect'}</Text>
          </Stack>
        ) : <Text>N/A</Text>
      )
    },
    {
      name: 'points',
      displayName: 'Points',
      minWidth: 80, // Adjusted width
      maxWidth: 100,
      // Ensure points values exist and are numbers
      render: (item) => <Text>{(typeof item?.earnedPoints === 'number' ? item.earnedPoints : '-')}/{(typeof item?.points === 'number' ? item.points : '-')}</Text>
    }
  ], []);

  // Render summary view
  const renderSummaryView = (): JSX.Element => {
    return (
      <div className={styles.scoreCard}>
        <div className={styles.scoreBadge} style={{
          backgroundColor: summaryPercentage >= 90 ? '#107c10' :
            summaryPercentage >= 70 ? '#498205' :
              summaryPercentage >= 50 ? '#ffaa44' : '#d13438'
        }}>
          <Text styles={scoreValueStyles}>{summaryPercentage}%</Text>
          <Text styles={scoreTextStyles}>Score</Text>
        </div>

        <Text styles={scoreDetailsStyles}>
          You answered {score} out of {totalPoints} points correctly.
        </Text>

        <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center" verticalAlign="center">
          <Icon iconName={getScoreIconName(summaryPercentage)} styles={{ root: { fontSize: '20px' } }} />
          <Text styles={resultMessageStyles}>{resultMessage}</Text>
        </Stack>

        <Stack horizontalAlign="center" tokens={stackTokens}>
          <ProgressIndicator
            percentComplete={summaryPercentage / 100}
            styles={{
              root: { maxWidth: '400px', margin: '20px auto 0' }, // Added top margin
              itemName: { textAlign: 'left' },
              itemProgress: { paddingBottom: 4 }
            }}
          />
        </Stack>
      </div>
    );
  };

  // Render detailed results view
  const renderDetailedView = (): JSX.Element => {
    // Add checks for detailedResults and its properties
    if (!detailedResults || !detailedResults.questionResults) {
      return (
        <MessageBar messageBarType={MessageBarType.info}>
          Detailed results are not available or loading.
        </MessageBar>
      );
    }


    return (
      <div className={styles.resultsContainer}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: '16px' }}>
          <Stack>
            <Text variant="large" style={{ fontWeight: FontWeights.semibold, color: '#0078d4' }}>
              {detailedResults.score} / {detailedResults.totalPoints} points ({detailedResults.percentage}%)
            </Text>
            {/* Re-added call to getPerformanceMessage */}
            <Text variant="medium">
              {getPerformanceMessage(detailedResults.percentage)}
            </Text>
          </Stack>
        </Stack>

        <Separator>Question Results</Separator>

        {/* Use ListView for a summary table */}
        <ListView
          items={detailedResults.questionResults}
          viewFields={viewFields}
          compact={false}
          selectionMode={SelectionMode.none}
          showFilter={false}
        />

        <Separator>Question Details</Separator>

        {/* Map over questions and render an Accordion for each */}
        {/* Use styles.questionAccordion or define questionAccordionContainer in SCSS */}
        <div className={styles.questionAccordion}>
          {detailedResults.questionResults.map((question, index) => {
            // Check if question object is valid
            if (!question) return null;

            // Create a simple string title for the Accordion header
            const simpleAccordionTitle = `Question ${index + 1}: ${question.title || 'Untitled Question'}`;

            // Construct the content JSX for this specific question's Accordion
            const accordionContent = (
              <Stack tokens={stackTokens} className={styles.questionDetails}>
                {/* Display Correct/Incorrect Icon + Points INSIDE the content */}
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: '10px' }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <Icon
                      iconName={question.isCorrect ? 'CheckMark' : 'Cancel'}
                      className={question.isCorrect ? styles.correctIcon : styles.incorrectIcon}
                    />
                    <Text variant="mediumPlus" style={{ fontWeight: FontWeights.semibold }}>
                      Result: {question.isCorrect ? 'Correct' : 'Incorrect'}
                    </Text>
                  </Stack>
                  <Text>Points: {question.earnedPoints} / {question.points}</Text>
                </Stack>
                <Separator />

                <Stack horizontal tokens={{ childrenGap: 20 }} wrap style={{ marginTop: '10px' }}>
                  <Stack className={styles.detailColumn}>
                    <Label>Your Answer:</Label>
                    <div className={styles.answerBox}>
                      { }
                      <Text>{formatAnswer(question.userAnswer)}</Text>
                    </div>
                  </Stack>

                  <Stack className={styles.detailColumn}>
                    <Label>Correct Answer:</Label>
                    <div className={styles.correctAnswerBox}>
                      { }
                      <Text>{formatAnswer(question.correctAnswer)}</Text>
                    </div>
                  </Stack>
                </Stack>

                {question.explanation && (
                  <MessageBar
                    messageBarType={MessageBarType.info}
                    className={styles.explanationBox} // Ensure this class exists in SCSS
                    styles={{ root: { marginTop: '16px' } }} // Add some top margin
                  >
                    <Label>Explanation:</Label>
                    <Text>{question.explanation}</Text>
                  </MessageBar>
                )}
              </Stack>
            );

            return (
              <Accordion
                key={question.id || index}
                title={simpleAccordionTitle}
                defaultCollapsed={true}
                className={styles.questionAccordion}
              >
                {/* The content goes here, as children */}
                {accordionContent}
              </Accordion>
            );
          })}
        </div>
      </div>
    );
  };

  // Main return statement
  return (
    <div className={styles.resultsContainer}>
      <Text styles={resultsTitleStyles}>Quiz Results</Text>

      {isSubmitting ? (
        <Stack horizontalAlign="center" tokens={stackTokens} style={{ padding: '30px' }}>
          <Spinner size={SpinnerSize.large} label="Submitting your results..." />
        </Stack>
      ) : (
        <>
          {submissionSuccess && (
            <MessageBar
              messageBarType={MessageBarType.success}
              isMultiline={false}
              styles={{ root: { marginBottom: '20px' } }}
            >
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                <Icon iconName={checkIconName} />
                <span>{messages.success || 'Your score has been successfully recorded!'}</span>
              </Stack>
            </MessageBar>
          )}

          {submissionError && (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true} // Allow multiline for potentially longer errors
              styles={{ root: { marginBottom: '20px' } }}
            >
              {submissionError} {/* Display the actual error message */}
            </MessageBar>
          )}

          {/* Show Pivot only if detailed results are available */}
          {detailedResults ? (
            <Pivot
              selectedKey={activeView}
              onLinkClick={(item) => item && setActiveView(item.props.itemKey || 'summary')}
              styles={{ root: { marginBottom: '20px' } }}
              headersOnly={false} // Ensure content area is shown
            >
              <PivotItem headerText="Summary" itemKey="summary">
                {/* Render summary view content here directly if Pivot handles content */}
                {renderSummaryView()}
              </PivotItem>
              <PivotItem headerText="Detailed Results" itemKey="details">
                {/* Render detailed view content here directly */}
                {renderDetailedView()}
              </PivotItem>
            </Pivot>
          ) : (
            // If no detailed results, just show summary
            renderSummaryView()
          )}

          {/* Conditionally render buttons based on view or always show */}
          <Stack horizontal horizontalAlign="center" tokens={stackTokens} style={{ marginTop: '20px' }}>
            <PrimaryButton
              text="Retake Quiz"
              onClick={onRetakeQuiz}
              iconProps={retakeIcon}
              styles={{ root: { minWidth: '140px' } }}
            />
            {/* Logic for showing "View Detailed Results" button can be simplified */}
            {/* Show if detailed results exist and current view is summary */}
            {activeView === 'summary' && detailedResults && (
              <DefaultButton
                text="View Detailed Results"
                onClick={() => setActiveView('details')}
                iconProps={reviewIcon}
              />
            )}
          </Stack>
        </>
      )}
    </div>
  );
};

export default QuizResults;