import * as React from 'react';
import { IQuizResultsProps } from './interfaces';
import {
  Stack,
  IStackTokens,
  Text,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  mergeStyles,
  Icon,
  ITextStyles,
  FontWeights,
  IIconProps
} from '@fluentui/react';

// Icons - use literal strings for iconName instead of possibly undefined values
const retakeIcon: IIconProps = { iconName: 'Refresh' };
const excellentIconName = 'Trophy';
const goodIconName = 'Emoji2';
const averageIconName = 'EmojiNeutral';
const poorIconName = 'Sad';
const checkIconName = 'CheckMark';

// Styles
const resultsContainerClass = mergeStyles({
  maxWidth: '800px',
  margin: '0 auto',
  padding: '20px',
  textAlign: 'center'
});

const resultsTitleStyles: ITextStyles = {
  root: {
    fontSize: '24px',
    fontWeight: FontWeights.semibold,
    marginBottom: '20px'
  }
};

const scoreCardClass = mergeStyles({
  background: 'white',
  padding: '30px 20px',
  borderRadius: '4px',
  boxShadow: '0 2px 8px rgba(0, 0, 0, 0.1)',
  marginTop: '20px',
  marginBottom: '30px',
  textAlign: 'center',
  border: '1px solid #edebe9'
});

const scoreBadgeClass = mergeStyles({
  width: '160px',
  height: '160px',
  borderRadius: '80px',
  margin: '0 auto 24px auto',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  boxShadow: '0 0 15px rgba(0, 0, 0, 0.1)'
});

// Dynamic classes based on score
const getScoreBadgeClass = (percentage: number): string => {
  let backgroundColor = '#0078d4'; // Default blue
  
  if (percentage >= 90) {
    backgroundColor = '#107c10'; // Green for excellent
  } else if (percentage >= 70) {
    backgroundColor = '#498205'; // Light green for good
  } else if (percentage >= 50) {
    backgroundColor = '#ffaa44'; // Orange for average
  } else {
    backgroundColor = '#d13438'; // Red for poor
  }
  
  return mergeStyles([
    scoreBadgeClass,
    {
      backgroundColor
    }
  ]);
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
    messages
  } = props;
  
  const percentage = totalQuestions > 0 ? Math.round((score / totalPoints) * 100) : 0;
  
  // Determine result message based on percentage
  let resultMessage = '';
  if (percentage >= 90) {
    resultMessage = messages.excellent || 'Excellent! You have mastered this topic!';
  } else if (percentage >= 70) {
    resultMessage = messages.good || 'Good job! You have a solid understanding.';
  } else if (percentage >= 50) {
    resultMessage = messages.average || 'Not bad. There\'s room for improvement.';
  } else {
    resultMessage = messages.poor || 'Keep studying. You\'ll get better with practice.';
  }

  // Stack tokens for spacing
  const stackTokens: IStackTokens = {
    childrenGap: 16
  };
  
  return (
    <div className={resultsContainerClass}>
      <Text styles={resultsTitleStyles}>Quiz Results</Text>
      
      {isSubmitting ? (
        <Stack horizontalAlign="center" tokens={stackTokens}>
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
              isMultiline={false}
              styles={{ root: { marginBottom: '20px' } }}
            >
              {submissionError}
            </MessageBar>
          )}
          
          <div className={scoreCardClass}>
            <div className={getScoreBadgeClass(percentage)}>
              <Text styles={scoreValueStyles}>{percentage}%</Text>
              <Text styles={scoreTextStyles}>Score</Text>
            </div>
            
            <Text styles={scoreDetailsStyles}>
              You answered {score} out of {totalPoints} points correctly.
            </Text>
            
            <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center">
              <Icon iconName={getScoreIconName(percentage)} />
              <Text styles={resultMessageStyles}>{resultMessage}</Text>
            </Stack>
            
            <Stack horizontalAlign="center" tokens={stackTokens}>
              <ProgressIndicator 
                percentComplete={percentage / 100} 
                styles={{ 
                  root: { maxWidth: '400px', margin: '0 auto' },
                  itemName: { textAlign: 'left' },
                  itemProgress: { paddingBottom: 4 }
                }}
              />
            </Stack>
          </div>
          
          <Stack horizontalAlign="center" tokens={stackTokens}>
            <PrimaryButton
              text="Retake Quiz"
              onClick={onRetakeQuiz}
              iconProps={retakeIcon}
              styles={{ root: { minWidth: '140px' } }}
            />
          </Stack>
        </>
      )}
    </div>
  );
};

export default QuizResults;