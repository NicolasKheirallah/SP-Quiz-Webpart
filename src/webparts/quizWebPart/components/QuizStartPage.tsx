import * as React from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  Image,
  ImageFit,
  Separator,
  IStackTokens,
  FontIcon,
  DefaultButton,
  Panel,
  PanelType,
  List
} from '@fluentui/react';
import styles from './Quiz.module.scss';

export interface IQuizStartPageProps {
  title: string;
  onStartQuiz: () => void;
  totalQuestions: number;
  totalPoints: number;
  categories: string[];
  timeLimit?: number; // in seconds
  passingScore?: number; // percentage
  quizImage?: string; // optional image URL
  description?: string;
  hasSavedProgress?: boolean; // Add this prop to indicate saved progress
  onResumeQuiz?: () => void; // Add this prop for resuming
}

// Stack tokens
const stackTokens: IStackTokens = {
  childrenGap: 15
};

// Heading stack tokens with more spacing
const headingStackTokens: IStackTokens = {
  childrenGap: 25
};

// Feature stack tokens
const featureStackTokens: IStackTokens = {
  childrenGap: 10
};

const QuizStartPage: React.FC<IQuizStartPageProps> = (props) => {
  const {
    title,
    onStartQuiz,
    totalQuestions,
    totalPoints,
    categories,
    timeLimit,
    passingScore,
    quizImage,
    description,
    hasSavedProgress,
    onResumeQuiz
  } = props;

  // Format time from seconds to minutes/hours
  const formatTime = (seconds?: number): string => {
    if (!seconds) return 'No time limit';
    
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    
    if (hours > 0) {
      return `${hours} hour${hours !== 1 ? 's' : ''} ${minutes > 0 ? `${minutes} minute${minutes !== 1 ? 's' : ''}` : ''}`;
    }
    
    if (minutes > 0) {
      return `${minutes} minute${minutes !== 1 ? 's' : ''}`;
    }
    
    return `${seconds} seconds`;
  };
  
  // Unique categories (no duplicates)
  const uniqueCategories = [...new Set(categories)].filter(cat => cat !== 'All');
  
  // Instructions panel state
  const [isInstructionsPanelOpen, setIsInstructionsPanelOpen] = React.useState(false);
  
  return (
    <div className={styles.quizStartPage}>
      <Stack tokens={headingStackTokens} horizontalAlign="center">
        <Stack.Item align="center" className={styles.quizStartTitle}>
          <Text variant="xxLarge">{title}</Text>
        </Stack.Item>
        
        {quizImage && (
          <Image 
            src={quizImage} 
            alt={title}
            className={styles.quizStartImage}
            imageFit={ImageFit.contain}
            width={300}
            height={180}
          />
        )}
        
        {description && (
          <Stack.Item className={styles.quizDescription}>
            <Text>{description}</Text>
          </Stack.Item>
        )}
      </Stack>
      
      <Separator className={styles.separator}>Quiz Details</Separator>
      
      <Stack tokens={stackTokens} horizontal wrap horizontalAlign="center" className={styles.quizStatsContainer}>
        <Stack className={styles.quizStatItem} horizontalAlign="center" tokens={featureStackTokens}>
          <FontIcon iconName="QuizNew" className={styles.statIcon} />
          <Text variant="large">{totalQuestions}</Text>
          <Text>Questions</Text>
        </Stack>
        
        <Stack className={styles.quizStatItem} horizontalAlign="center" tokens={featureStackTokens}>
          <FontIcon iconName="Trophy" className={styles.statIcon} />
          <Text variant="large">{totalPoints}</Text>
          <Text>Total Points</Text>
        </Stack>
        
        <Stack className={styles.quizStatItem} horizontalAlign="center" tokens={featureStackTokens}>
          <FontIcon iconName="Timer" className={styles.statIcon} />
          <Text variant="large">{formatTime(timeLimit)}</Text>
          <Text>Time Limit</Text>
        </Stack>
        
        {passingScore && (
          <Stack className={styles.quizStatItem} horizontalAlign="center" tokens={featureStackTokens}>
            <FontIcon iconName="Ribbon" className={styles.statIcon} />
            <Text variant="large">{passingScore}%</Text>
            <Text>Passing Score</Text>
          </Stack>
        )}
        
        {uniqueCategories.length > 0 && (
          <Stack className={styles.quizStatItem} horizontalAlign="center" tokens={featureStackTokens}>
            <FontIcon iconName="BulletedList" className={styles.statIcon} />
            <Text variant="large">{uniqueCategories.length}</Text>
            <Text>Categories</Text>
          </Stack>
        )}
      </Stack>
      
      <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} className={styles.quizStartActions}>
        <PrimaryButton
          text="Start Quiz"
          iconProps={{ iconName: 'Play' }}
          onClick={onStartQuiz}
          className={styles.startButton}
        />
        
        {hasSavedProgress && onResumeQuiz && (
          <DefaultButton
            text="Resume Saved Quiz"
            iconProps={{ iconName: 'SkypeCircleCheck' }}
            onClick={onResumeQuiz}
            className={styles.resumeButton}
          />
        )}
        
        <DefaultButton
          text="Quiz Instructions"
          iconProps={{ iconName: 'ReadingMode' }}
          onClick={() => setIsInstructionsPanelOpen(true)}
        />
      </Stack>
      
      {/* Instructions Panel */}
      <Panel
        isOpen={isInstructionsPanelOpen}
        onDismiss={() => setIsInstructionsPanelOpen(false)}
        headerText="Quiz Instructions"
        closeButtonAriaLabel="Close"
        type={PanelType.medium}
      >
        <div className={styles.instructionsPanel}>
          <Text variant="large">How to Take This Quiz</Text>
          
          <Separator />
          
          <List
            items={[
              "Read each question carefully before selecting your answer.",
              `You have ${formatTime(timeLimit)} to complete this quiz.`,
              "Some questions may have individual time limits.",
              passingScore ? `You need to score at least ${passingScore}% to pass.` : "Try to answer all questions for the best score.",
              "Once you submit your answer for a question, you cannot change it.",
              "Click 'Submit Quiz' when you have completed all questions.",
              "Your results will be displayed at the end of the quiz.",
              "You can save your progress and resume later if needed."
            ]}
            onRenderCell={(item) => (
              <div className={styles.instructionItem}>
                <FontIcon iconName="CircleRing" className={styles.instructionIcon} />
                <Text>{item}</Text>
              </div>
            )}
          />
          
          <Separator />
          
          <Text variant="large">Categories in this Quiz</Text>
          <div className={styles.categoriesList}>
            {uniqueCategories.map((category, index) => (
              <div key={index} className={styles.categoryTag}>
                {category}
              </div>
            ))}
          </div>
          
          <div className={styles.startButtonContainer}>
            <PrimaryButton
              text="Start Quiz"
              iconProps={{ iconName: 'Play' }}
              onClick={() => {
                setIsInstructionsPanelOpen(false);
                onStartQuiz();
              }}
              className={styles.startButton}
            />
          </div>
        </div>
      </Panel>
    </div>
  );
};

export default QuizStartPage;