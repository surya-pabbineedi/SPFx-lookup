import * as redux from 'redux';
import { listReducer } from './listReducer';

export const rootReducer = redux.combineReducers({
  listReducer
});
