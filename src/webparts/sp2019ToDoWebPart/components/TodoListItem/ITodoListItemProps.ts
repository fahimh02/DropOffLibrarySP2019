import ITodoItem from '../../models/ITodoItem';
import ItemEditCallback from '../../models/ItemEditCallback';
import ItemOperationCallback from '../../models/ItemOperationCallback';

interface ITodoListItemProps {
  item: ITodoItem;
  isChecked?: boolean;
  onEditListItem: ItemEditCallback;
  onDeleteListItem: ItemOperationCallback;
}

export default ITodoListItemProps;