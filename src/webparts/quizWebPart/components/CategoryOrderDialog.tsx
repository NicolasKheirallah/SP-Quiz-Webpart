import * as React from 'react';
import { 
  Dialog, 
  DialogType, 
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  IDragDropEvents,
  IDragDropContext,
  IconButton,
  Stack,
  Text,
  IIconProps,
  Dropdown,
  IDropdownOption
} from '@fluentui/react';
import styles from './Quiz.module.scss';
import { IQuizQuestion } from './interfaces';

// Icons
const moveIcon: IIconProps = { iconName: 'Move' };

export interface ICategoryOrderDialogProps {
  categories: string[]; 
  questions: IQuizQuestion[];
  onUpdateCategories: (newCategories: string[]) => void;
  onDismiss: () => void;
  isOpen: boolean;
}
  
export interface ICategoryOrderDialogState {
  categories: ICategoryItem[];
  isEditMode: boolean;
  sortMethod: string;
}

interface ICategoryItem {
  key: string;
  name: string;
  count: number;
}

export default class CategoryOrderDialog extends React.Component<ICategoryOrderDialogProps, ICategoryOrderDialogState> {
  private _draggedItem: ICategoryItem | undefined;

  constructor(props: ICategoryOrderDialogProps) {
    super(props);

    // Initialize the state with categories from props.
    const categories: ICategoryItem[] = this.props.categories
      .filter(cat => cat !== 'All') // Exclude 'All' category from reordering.
      .map(cat => ({
        key: cat,
        name: cat,
        count: 0 // Initialize count to 0; will be updated in componentDidMount.
      }));

    this.state = {
      categories,
      isEditMode: true,
      sortMethod: 'custom'
    };
  }

  public componentDidMount(): void {
    this.updateCategoryCounts();
  }
  
  private updateCategoryCounts(): void {
    const { questions } = this.props;
    if (!questions || questions.length === 0) {
      return;
    }
    
    // Create a map to count questions per category.
    const categoryCounts = new Map<string, number>();
    questions.forEach(question => {
      if (question.category) {
        const count = categoryCounts.get(question.category) || 0;
        categoryCounts.set(question.category, count + 1);
      }
    });
    
    // Update category counts in state.
    const updatedCategories = this.state.categories.map(category => ({
      ...category,
      count: categoryCounts.get(category.name) || 0
    }));
    this.setState({ categories: updatedCategories });
  }
  
  public render(): React.ReactElement<ICategoryOrderDialogProps> {
    const { isOpen, onDismiss } = this.props;
    const { categories, isEditMode, sortMethod } = this.state;

    const columns: IColumn[] = [
      {
        key: 'move',
        name: '',
        className: styles.fileDragIcon,
        iconName: 'GripperBarHorizontal',
        iconClassName: styles.fileIconHeaderIcon,
        minWidth: 16,
        maxWidth: 16,
        isIconOnly: true,
        onRender: (item: ICategoryItem) => (
          <IconButton iconProps={moveIcon} title="Drag to reorder" disabled={!isEditMode} />
        )
      },
      {
        key: 'name',
        name: 'Category',
        fieldName: 'name',
        minWidth: 300,
        isRowHeader: true,
        data: 'string',
        onRender: (item: ICategoryItem) => <Text>{item.name}</Text>
      },
      {
        key: 'count',
        name: 'Questions',
        fieldName: 'count',
        minWidth: 70,
        maxWidth: 100,
        data: 'number',
        onRender: (item: ICategoryItem) => <Text>{item.count}</Text>
      }
    ];

    const sortOptions: IDropdownOption[] = [
      { key: 'custom', text: 'Custom Order' },
      { key: 'alphabetical', text: 'Alphabetical (A-Z)' },
      { key: 'alphabeticalReverse', text: 'Alphabetical (Z-A)' },
      { key: 'count', text: 'By Question Count' }
    ];

    // Update dragDropEvents with explicit parameter types instead of any.
    const dragDropEvents: IDragDropEvents = {
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext): boolean => true,
      canDrag: (item?: ICategoryItem): boolean => isEditMode,
      onDragEnter: (item?: ICategoryItem, event?: DragEvent): string => {
        return styles.categoryItemDragEnter;
      },
      onDragLeave: (item?: ICategoryItem, event?: DragEvent): void => {},
      onDrop: (item?: ICategoryItem, event?: DragEvent): void => {
        if (this._draggedItem && item) {
          this._insertBeforeItem(item);
        }
      },
      onDragStart: (
        item?: ICategoryItem,
        itemIndex?: number,
        selectedItems?: ICategoryItem[],
        event?: MouseEvent
      ): void => {
        this._draggedItem = item;
      },
      onDragEnd: (item?: ICategoryItem, event?: DragEvent): void => {
        this._draggedItem = undefined;
      }
    };

    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={onDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Manage Category Order',
          subText: 'Drag and drop categories to change their order, or use a predefined sorting method.'
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 500 } }
        }}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <Dropdown
            label="Sort Method"
            selectedKey={sortMethod}
            options={sortOptions}
            onChange={this._onSortMethodChange}
          />

          <DetailsList
            items={categories}
            columns={columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            isHeaderVisible={true}
            dragDropEvents={dragDropEvents}
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton onClick={this._applyChanges} text="Apply" />
          <DefaultButton onClick={onDismiss} text="Cancel" />
        </DialogFooter>
      </Dialog>
    );
  }

  private _onSortMethodChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    
    const sortMethod = option.key as string;
    const sortedCategories = [...this.state.categories]; // 'const' used because it is not reassigned.
    
    switch (sortMethod) {
      case 'alphabetical':
        sortedCategories.sort((a, b) => a.name.localeCompare(b.name));
        break;
      case 'alphabeticalReverse':
        sortedCategories.sort((a, b) => b.name.localeCompare(a.name));
        break;
      case 'count':
        sortedCategories.sort((a, b) => b.count - a.count);
        break;
      // 'custom' maintains the current order.
    }
    
    this.setState({ 
      categories: sortedCategories,
      sortMethod,
      isEditMode: sortMethod === 'custom'
    });
  };

  private _insertBeforeItem = (item: ICategoryItem): void => {
    const draggedItems = [this._draggedItem!];
    const insertIndex = this.state.categories.indexOf(item);
    const items = this.state.categories.filter(
      (category: ICategoryItem) => draggedItems.indexOf(category) === -1
    );
    items.splice(insertIndex, 0, ...draggedItems);
    this.setState({ categories: items });
  };

  private _applyChanges = (): void => {
    // Convert category items back to a string array and prepend 'All'.
    const newCategories = ['All', ...this.state.categories.map(item => item.name)];
    this.props.onUpdateCategories(newCategories);
    this.props.onDismiss();
  };
}
