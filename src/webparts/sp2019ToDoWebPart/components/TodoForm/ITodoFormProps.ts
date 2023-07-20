import ITodoTaskList from '../../models/ITodoTaskList';
import ItemCreationCallback from '../../models/ItemCreationCallback';

interface ITodoFormProps {
  onAddTodoItem: ItemCreationCallback;
  list: ITodoTaskList[];
}

export default ITodoFormProps;