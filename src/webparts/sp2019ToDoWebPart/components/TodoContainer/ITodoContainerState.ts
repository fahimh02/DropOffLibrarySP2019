import ITodoItem from '../../models/ITodoItem';
import ITodoTaskList from '../../models/ITodoTaskList';

interface ITodoContainerState {
  todoItems: ITodoItem[];
  libraries:ITodoTaskList[];
  isLoading:boolean;
  showDialog:boolean;
}

export default ITodoContainerState;