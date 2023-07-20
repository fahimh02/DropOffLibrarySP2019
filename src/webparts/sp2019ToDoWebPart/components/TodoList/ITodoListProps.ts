import ITodoItem from '../../models/ITodoItem';
import ItemEditCallback from '../../models/ItemEditCallback';
import ItemOperationCallback from '../../models/ItemOperationCallback';

interface ITodoListProps {
  items: ITodoItem[];
  onEditTodoItem: ItemEditCallback;
  onDeleteTodoItem: ItemOperationCallback;
}

export default ITodoListProps;