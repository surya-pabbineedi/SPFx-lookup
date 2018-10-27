export const listReducer = (state = [], action) => {
  switch (action.type) {
    case 'LIST_DATA':
      return [...state, (<any>Object).assign({}, action.data)];
    default:
      return state;
  }
};
