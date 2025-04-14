import * as React from 'react';
import {
  Stack,
  Dropdown,
  IDropdownOption,
  Label,
  IStackTokens
} from '@fluentui/react';
import { IMatchingPair } from './interfaces';
import styles from './Quiz.module.scss';

export interface IMatchingQuestionProps {
  questionId: number;
  matchingPairs: IMatchingPair[];
  onMatchingPairsChange: (questionId: number, updatedPairs: IMatchingPair[]) => void;
  disabled?: boolean;
}

// Stack tokens
const stackTokens: IStackTokens = {
  childrenGap: 12
};

export class MatchingQuestion extends React.Component<IMatchingQuestionProps, { currentPairs: IMatchingPair[] }> {
  
  constructor(props: IMatchingQuestionProps) {
    super(props);
    
    this.state = {
      currentPairs: [...props.matchingPairs]
    };
  }
  
  // Update state when props change
  public componentDidUpdate(prevProps: IMatchingQuestionProps): void {
    if (prevProps.matchingPairs !== this.props.matchingPairs) {
      this.setState({ currentPairs: [...this.props.matchingPairs] });
    }
  }
  
  // Generate options for the right-side items dropdown
  private generateRightItemOptions = (): IDropdownOption[] => {
    return this.props.matchingPairs.map(pair => ({
      key: pair.id,
      text: pair.rightItem
    }));
  }
  
  // Handle when a user selects a right-side item for a left-side item
  private handleMatchSelection = (pairId: string, selectedRightId?: string): void => {
    // Update the current pairs with the user's selection
    const updatedPairs = this.state.currentPairs.map(pair => 
      pair.id === pairId 
        ? { ...pair, userSelectedRightId: selectedRightId } 
        : pair
    );
    
    this.setState({ currentPairs: updatedPairs });
    
    // Notify parent component about the change
    this.props.onMatchingPairsChange(this.props.questionId, updatedPairs);
  }
  
  public render(): React.ReactElement<IMatchingQuestionProps> {
    const { disabled } = this.props;
    const { currentPairs } = this.state;
    
    return (
      <Stack tokens={stackTokens} className={styles.matchingQuestion}>
        {currentPairs.map((pair) => (
          <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center" key={pair.id} className={styles.matchingRow}>
            <Stack.Item className={styles.matchingLeftItem}>
              <Label>{pair.leftItem}</Label>
            </Stack.Item>
            
            <Stack.Item className={styles.matchingRightDropdown}>
              <Dropdown
                placeholder="Select the matching item"
                selectedKey={pair.userSelectedRightId}
                options={this.generateRightItemOptions()}
                onChange={(_, option) => this.handleMatchSelection(pair.id, option?.key as string)}
                disabled={disabled}
              />
            </Stack.Item>
          </Stack>
        ))}
      </Stack>
    );
  }
}

export default MatchingQuestion;