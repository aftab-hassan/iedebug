import DeleteType from '@ms/odsp-datasources/lib/dataSources/recycleBin/DeletionType';
import { ISPDeleteItemsContext } from '@ms/odsp-datasources/lib/dataSources/recycleBin/ISPDeleteItemsContext';
import { RecycleBinDataSource } from '@ms/odsp-datasources/lib/dataSources/recycleBin/RecycleBinDataSource';
import ISpPageContext from '@ms/odsp-datasources/lib/interfaces/ISpPageContext';
import { IListItemFormUpdateValue } from '@ms/odsp-datasources/lib/models/clientForm/IClientForm';
import { ColumnFieldType, ISPListItem } from '@ms/odsp-datasources/lib/SPListItemProcessor';
import { ISPListRow } from '@ms/odsp-datasources/lib/SPListItemRetriever';
import { createResolvedResourceKey, ResourceKey } from '@ms/odsp-utilities/lib/resources/Resources';
import {
  IListItemUpdateResult,
  SPListItemDataSource,
  IListItemUpdate
} from '../../library/api/item/SPListItemDataSource';
import { resourceKey as listItemDataSourceKey } from '../../library/api/item/SPListItemDataSource.key';
import { IField } from '../../library/api/list/ISPField';
import { pageContextKey } from '../../resources/ResourceKeys';
import { getItemKeyResourceKey } from '../../stores/list/GetItemKey.key';
import { listItemSelectionStoreKey } from '../../stores/list/ListItemSelectionStore.key';
import {
  ListItemStore,
  listItemStoreKey,
  IListItemStatusUpdateParams,
  IListItemStatus,
  ListItemMessageMap
} from '../../stores/list/ListItemStore.key';
import { IEditor } from '@ms/odsp-datasources/lib/SPListItemRetriever';
import { GetHelpers } from '../../library/api/core/HelpersBootstrap';
import { IHelpers } from '../../library/api/core/IHelpers';
import { AddNewRowIdPrefix } from '../../library/controls/list/ListHelper';
import { createNewRowItemPublisher, createNewListItemPublisher } from './ListItemProvider.publisher';

export interface IListItemProviderParams {}
export interface IListItemProviderDependencies {
  pageContext: ISpPageContext;
  listItemStore: ListItemStore;
  listItemSelectionStore: typeof listItemSelectionStoreKey.type;
  listItemDataSource: SPListItemDataSource;
  getItemKey: (item: ISPListRow) => string;
}

/**
 * Public methods that the ListItemProvider supports
 */
export interface IListItemProvider {
  updateListItemField: (
    item: ISPListRow,
    field: IField,
    newValue: any,
    allFields: IField[]
  ) => Promise<IListItemUpdateResult>;
  updateListItemBatch: (
    items: ISPListRow[],
    fields: IField[],
    allFields: IField[]
  ) => Promise<IListItemUpdateResult[]>;
  deleteItems: (items: ISPListRow[], resultCallback: (result: any) => void) => void;
}

export type IFieldError = { [fieldName: string]: string };
export interface IFieldErrors {
  errors?: IFieldError;
}
export type ISPListRowWithErrors = ISPListRow & IFieldErrors;

interface INotifyStoreOfUpdatesParams {
  itemKey: string;
  updatedFields: IListItemFormUpdateValue[];
  listRow: ISPListRow;
  valueBeforeSave?: any;
  field?: IField;
}

export class ListItemProvider implements IListItemProvider {
  private _pageContext: ISpPageContext;
  private _dataSource: SPListItemDataSource;
  private _listItemStore: ListItemStore;
  private _listItemSelectionStore: typeof listItemSelectionStoreKey.type;
  private _getItemKey: (item: ISPListRow) => string;

  constructor(params: IListItemProviderParams, dependencies: IListItemProviderDependencies) {
    const { pageContext, listItemStore, listItemSelectionStore, listItemDataSource, getItemKey } =
      dependencies;

    this._pageContext = pageContext;
    this._listItemStore = listItemStore;
    this._listItemSelectionStore = listItemSelectionStore;
    this._dataSource = listItemDataSource;
    this._getItemKey = getItemKey;
  }

  private _wait(milliseconds) {
    return new Promise((resolve) => setTimeout(resolve, milliseconds));
  }

  /**
   * @param item
   * @param targetField Field being edited
   * @param newValue is of type ISPFieldValue. Figure out how to make this not be any.
   * @param allFields All fields of the item. If an error exists on the item, all fields will be submitted in payload. This
   * is to ensure that if multiple fields have a validation error, updating only one field will not effectively remove
   * the other errors -- all fields should be submitted for validation
   */
  public updateListItemField(item: ISPListRow, targetField: IField, newValue: any, allFields: IField[]) {
    // If newValue has a rawValue prop, try to use that first. That is what
    // validateUpdateListItem API expects.
    const rawValue = newValue.rawValue || newValue.value;

    const { SPListHelpers }: IHelpers = GetHelpers();
    const itemKey = this._listItemStore.getItemKey(item);

    const oldStatus = this._listItemStore?.getItemStatus(itemKey);
    const hasError = oldStatus?.hasError;

    let shouldMakeServerRequest = true; // whether we want to make a call to the server
    let useAllFields = false; // whether we want to update all fields

    // If an error exists on the item, don't send call to server to validate fields
    // unless user is editing last field in fieldWithErrors
    if (hasError) {
      shouldMakeServerRequest = false;
      const fieldsWithErrors = this._listItemStore.getItemStatus(itemKey)?.fieldsWithErrors;

      // check if targetField is last error on item
      // if so, we want to resubmit and validate all fields
      if (Object.keys(fieldsWithErrors).length === 1 && fieldsWithErrors[targetField.realFieldName]) {
        useAllFields = true;
        shouldMakeServerRequest = true;
      }
    }

    const formValues = (useAllFields &&
      allFields.map((currentField: IField) => {
        let fieldValue = item[currentField.realFieldName];

        if (targetField.realFieldName === currentField.realFieldName) {
          fieldValue = rawValue;
        } else if (fieldValue) {
          if (currentField.type === ColumnFieldType.User && Array.isArray(fieldValue)) {
            fieldValue = fieldValue.map((user: IEditor & { Key: string }) => {
              user.Key = user.email;
              return user;
            });
            fieldValue = JSON.stringify(fieldValue);
          } else if (currentField.type === ColumnFieldType.Thumbnail && typeof fieldValue !== 'string') {
            fieldValue = JSON.stringify(fieldValue);
          } else if (currentField.type !== ColumnFieldType.DateTime) {
            const initialValue = SPListHelpers.getSPFieldValue(currentField, item);
            fieldValue = initialValue.rawValue || initialValue.value;
          }
          // else, currentField is a DateTime which does not need to be
          // modified in this case
        }

        return {
          FieldName: currentField.realFieldName,
          FieldValue: fieldValue,
          HasException: false,
          ErrorMessage: ''
        };
      })) || [
      {
        FieldName: targetField.realFieldName,
        FieldValue: rawValue,
        HasException: false,
        ErrorMessage: ''
      }
    ];

    if (shouldMakeServerRequest) {
      const oldStatus = this._listItemStore?.getItemStatus(itemKey);
      this._listItemStore.updateItemsStatus('ListItemProvider.updateListItemField.updateItemsStatus', [
        { itemKey, status: { ...oldStatus, isUpdating: true } }
      ]);

      // return this._wait(5000).then(() => {
      //   return this._dataSource
      //     .validateUpdateListItem(
      //       this._pageContext.listUrl!,
      //       Number(item.ID),
      //       formValues,
      //       false /*bNewDocumentUpdate*/,
      //       null /*checkInComment*/
      //       // We no longer send this value because setting it to true
      //       // renders the date in a way that returns a server response of 500
      //       // rather than 200 with item.HasException for an invalid date
      //       // true /*datesInUTC*/
      //     )
      //     .then((updatedFields: IListItemUpdateResult) => {
      //       // At this point, the API has returned.
      //       // Tell the ItemStore that some items have changed
      //       const itemKey = this._listItemStore.getItemKey(item);
      //       this._notifyItemStoreOfFieldUpdates([
      //         {
      //           itemKey,
      //           updatedFields: updatedFields.listFormValues,
      //           listRow: updatedFields.listRow,
      //           valueBeforeSave: newValue,
      //           field: targetField
      //         }
      //       ]);
      //       return updatedFields;
      //     });
      // });

      return this._dataSource
        .validateUpdateListItem(
          this._pageContext.listUrl!,
          Number(item.ID),
          formValues,
          false /*bNewDocumentUpdate*/,
          null /*checkInComment*/
          // We no longer send this value because setting it to true
          // renders the date in a way that returns a server response of 500
          // rather than 200 with item.HasException for an invalid date
          // true /*datesInUTC*/
        )
        .then((updatedFields: IListItemUpdateResult) => {
          // At this point, the API has returned.
          // Tell the ItemStore that some items have changed
          const itemKey = this._listItemStore.getItemKey(item);
          this._notifyItemStoreOfFieldUpdates([
            {
              itemKey,
              updatedFields: updatedFields.listFormValues,
              listRow: updatedFields.listRow,
              valueBeforeSave: newValue,
              field: targetField
            }
          ]);
          return updatedFields;
        });
    } else {
      this._updateItemStore(item, targetField, newValue);
    }
  }

  // This function calls the backend API to create a new list item
  public createListItem(item: ISPListRow) {
    let formValues = new Array();

    const fieldsChanged = Object.keys(item);
    for (let i = 0; i < fieldsChanged.length; i++) {
      const fieldName = fieldsChanged[i];

      if (fieldName !== 'ID') {
        formValues.push({
          FieldName: fieldName,
          FieldValue: item[fieldName],
          ErrorMessage: 'null',
          HasException: false
        });
      }
    }

    const itemStatusesToUpdate: IListItemStatusUpdateParams<string, IListItemStatus>[] = [];
    itemStatusesToUpdate.push({
      itemKey: item.ID,
      status: { isUpdating: true }
    });
    this._listItemStore.updateItemsStatus('ListItemProvider', itemStatusesToUpdate, true);

    // return this._wait(Number(item.Title) % 2 === 0 ? 5000 : 10000).then(() => {
    //   return this._dataSource
    //     .validateCreateListItem(
    //       this._pageContext.listUrl!,
    //       Number(item.ID),
    //       formValues,
    //       false /*bNewDocumentUpdate*/,
    //       null /*checkInComment*/
    //       // We no longer send this value because setting it to true
    //       // renders the date in a way that returns a server response of 500
    //       // rather than 200 with item.HasException for an invalid date
    //       // true /*datesInUTC*/
    //     )
    //     .then((updatedFields: IListItemUpdateResult) => {
    //       // Insert the value returned from the server response
    //       // to ItemStore now that the row actually exists
    //       let idOfNewRowCreated;
    //       const listFormValuesReturnedFromServer = updatedFields.listFormValues;
    //       for (let i = 0; i < listFormValuesReturnedFromServer.length; i++) {
    //         const listFormValueReturnedFromServer = listFormValuesReturnedFromServer[i];
    //         if (listFormValueReturnedFromServer && listFormValueReturnedFromServer.FieldName === 'Id') {
    //           // Find the ID of the New Row that was just created
    //           idOfNewRowCreated = listFormValueReturnedFromServer.FieldValue;
    //           break;
    //         }
    //       }

    //       /*
    //       Make additional call to update row with an empty set [] of updates so as to get the server response in
    //       updatedFields.listRow in the format that the renderers of the field editors need
    //       to populate values
    //       - This is necessary because the AddValidateUpdateItemUsingPath API used above inside validateCreateListItem
    //       does not return the server response with all the fields that the field editors need to render.
    //       So, instead use the ValidateUpdateFetchListItem in validateUpdateListItem below to get the
    //       required response.
    //     */
    //       // TODO: (afhassan) Check if there is a GET API that returns server response in the format that
    //       // renderers need as opposed to using the ValidateUpdateFetchListItem API with an empty set of updates
    //       if (idOfNewRowCreated) {
    //         return this._dataSource
    //           .validateUpdateListItem(
    //             this._pageContext.listUrl!,
    //             Number(idOfNewRowCreated),
    //             [],
    //             false /*bNewDocumentUpdate*/,
    //             null /*checkInComment*/
    //             // We no longer send this value because setting it to true
    //             // renders the date in a way that returns a server response of 500
    //             // rather than 200 with item.HasException for an invalid date
    //             // true /*datesInUTC*/
    //           )
    //           .then((updatedFields: IListItemUpdateResult) => {
    //             this._listItemStore.addNewItems(
    //               item && item.ID.includes(AddNewRowIdPrefix)
    //                 ? createNewRowItemPublisher
    //                 : createNewListItemPublisher,
    //               [updatedFields.listRow]
    //             );
    //           });
    //       }
    //     });
    // });
    //
    return this._dataSource
      .validateCreateListItem(
        this._pageContext.listUrl!,
        Number(item.ID),
        formValues,
        false /*bNewDocumentUpdate*/,
        null /*checkInComment*/
        // We no longer send this value because setting it to true
        // renders the date in a way that returns a server response of 500
        // rather than 200 with item.HasException for an invalid date
        // true /*datesInUTC*/
      )
      .then((updatedFields: IListItemUpdateResult) => {
        // Insert the value returned from the server response
        // to ItemStore now that the row actually exists
        let idOfNewRowCreated;
        const listFormValuesReturnedFromServer = updatedFields.listFormValues;
        for (let i = 0; i < listFormValuesReturnedFromServer.length; i++) {
          const listFormValueReturnedFromServer = listFormValuesReturnedFromServer[i];
          if (listFormValueReturnedFromServer && listFormValueReturnedFromServer.FieldName === 'Id') {
            // Find the ID of the New Row that was just created
            idOfNewRowCreated = listFormValueReturnedFromServer.FieldValue;
            break;
          }
        }

        /*
          Make additional call to update row with an empty set [] of updates so as to get the server response in
          updatedFields.listRow in the format that the renderers of the field editors need
          to populate values
          - This is necessary because the AddValidateUpdateItemUsingPath API used above inside validateCreateListItem
          does not return the server response with all the fields that the field editors need to render.
          So, instead use the ValidateUpdateFetchListItem in validateUpdateListItem below to get the
          required response.
        */
        // TODO: (afhassan) Check if there is a GET API that returns server response in the format that
        // renderers need as opposed to using the ValidateUpdateFetchListItem API with an empty set of updates
        if (idOfNewRowCreated) {
          return this._dataSource
            .validateUpdateListItem(
              this._pageContext.listUrl!,
              Number(idOfNewRowCreated),
              [],
              false /*bNewDocumentUpdate*/,
              null /*checkInComment*/
              // We no longer send this value because setting it to true
              // renders the date in a way that returns a server response of 500
              // rather than 200 with item.HasException for an invalid date
              // true /*datesInUTC*/
            )
            .then((updatedFields: IListItemUpdateResult) => {
              this._listItemStore.addNewItems(
                item && item.ID.includes(AddNewRowIdPrefix)
                  ? createNewRowItemPublisher
                  : createNewListItemPublisher,
                [updatedFields.listRow]
              );
            });
        }
      });
  }

  /**
   * @param items
   * @param fields Fields being edited
   * @param allFields All fields of the item. If an error exists on the item, all fields will be submitted in payload. This
   * is to ensure that if multiple fields have a validation error, updating only one field will not effectively remove
   * the other errors -- all fields should be submitted for validation
   */
  public updateListItemBatch(items: ISPListRow[], fields: IField[], allFields: IField[]) {
    const itemsUpdating: IListItemStatusUpdateParams<string, IListItemStatus>[] = [];

    let shouldMakeServerRequest = true; // whether or not we want to make a call to the server

    const newItems: IListItemUpdate[] = items.map((item: ISPListRow) => {
      const oldStatus = this._listItemStore?.getItemStatus(item.ID); // item.ID == itemKey
      const hasError = oldStatus?.hasError;
      let formValuesFields = fields;

      const itemKey = this._listItemStore.getItemKey(item);

      // If an error exists on the item, don't send call to server to validate fields
      // unless user is editing last field in fieldWithErrors
      if (hasError) {
        shouldMakeServerRequest = false;
        const fieldsWithErrors = this._listItemStore.getItemStatus(itemKey)?.fieldsWithErrors;

        // check if remaining fieldsWithErrors are being edited by
        // batch fields, if so, we want to resubmit and validate all fields
        let numberOfEditedErrorFields = 0;
        for (const field of fields) {
          if (fieldsWithErrors[field.realFieldName]) {
            numberOfEditedErrorFields++;
          }
        }
        if (Object.keys(fieldsWithErrors).length === numberOfEditedErrorFields) {
          // Use allFields only when last field with error is edited, so that all fields get re-validated
          formValuesFields = allFields;
          shouldMakeServerRequest = true;
        }
      }

      return {
        itemId: item.ID,
        formValues: formValuesFields.map((field: IField) => {
          let fieldValue = item[field.realFieldName];

          if (!shouldMakeServerRequest) {
            this._updateItemStore(item, field, fieldValue);
          } else {
            itemsUpdating.push({ itemKey: item.ID, status: { ...oldStatus, isUpdating: true } });

            if (formValuesFields === allFields) {
              // `Key` property must be used in payload for resolving the the user server side
              if (field.type === ColumnFieldType.User && Array.isArray(fieldValue)) {
                fieldValue = fieldValue.map((user: IEditor & { Key: string }) => {
                  user.Key = user.email;
                  return user;
                });
              }
            }
          }
          if (field.type === ColumnFieldType.User && fieldValue) {
            fieldValue = JSON.stringify(fieldValue);
          }

          return {
            FieldName: field.realFieldName,
            FieldValue: fieldValue,
            HasException: false,
            ErrorMessage: ''
          };
        }),
        bNewDocumentUpdate: false,
        checkInComment: null
        // We no longer send this value because setting it to true
        // renders the date in a way that returns a server response of 500
        // rather than 200 with item.HasException for an invalid date
        // datesInUTC: true
      };
    });

    this._listItemStore.updateItemsStatus(
      'ListItemProvider.updateListItemBatch.updateItemsStatus',
      itemsUpdating
    );

    if (shouldMakeServerRequest) {
      return this._dataSource
        .validateUpdateListItemBatch(this._pageContext.listUrl!, newItems)
        .then((updatedFields: IListItemUpdateResult[]) => {
          const fieldUpdates = updatedFields.map((result: IListItemUpdateResult) => {
            return {
              itemKey: result.listRow.ID,
              updatedFields: result.listFormValues,
              listRow: result.listRow
            };
          });

          this._notifyItemStoreOfFieldUpdates(fieldUpdates);

          return updatedFields;
        });
      // TODO: Do something with response?
    }
  }

  /**
   * Calls the delete API
   * Invoke the recycleBin datasource in odsp-common to do the actual delete
   * @param items
   */
  public deleteItems(items: ISPListRow[], resultCallback: (deleteResults: any) => void) {
    let itemsToDelete: ISPListItem[] = items.map((item: ISPListRow) => {
      return { key: item.ID, properties: item };
    });

    let deleteContext: ISPDeleteItemsContext = {
      items: itemsToDelete,
      deletionType: DeleteType.Recycle,
      listId: this._pageContext.listId, //TODO: Is this right? Do we need a better list context?
      parentKey: '' //TODO: Is this right?
    };

    let recycleBinDS = new RecycleBinDataSource({ pageContext: this._pageContext });
    recycleBinDS
      .deleteItems(deleteContext) // TODO: pass in Engagement Executor
      .then((result: string[]) => {
        //on success
        const itemKeysDeleted: string[] = items.map((item: ISPListRow) => this._getItemKey(item)); //Get a list of item IDs (one for each item deleted)
        this._processDeletedItems(itemKeysDeleted);
        resultCallback(null /* No errors to process */); // call back the caller with the result
      })
      .catch((e: Error) => {
        this._handleDeleteErrors(e, resultCallback);
      });
  }

  private _processDeletedItems(itemKeysDeleted: string[]) {
    const publisher = 'ListItemProvider.deleteItems';
    this._listItemStore.deleteItems(publisher, itemKeysDeleted);
    this._listItemSelectionStore.remove(publisher, itemKeysDeleted);
  }

  /**
   * Handle the case when the API returns errors. We cycle through each
   * item in the result array looking for successful deletes. If we find them
   * we need to delete them from the list view.
   * @param result
   * @param resultCallback
   */
  private _handleDeleteErrors(result: any, resultCallback: (deleteResults: any) => void) {
    if (result && result.data && result.data.items) {
      const itemKeysDeleted: string[] = [];
      //Cycle through each returned item status
      for (let i = 0; i < result.data.items.length; i++) {
        let itemResult = result.data.items[i];
        if (!Boolean(itemResult.error)) {
          // Item was deleted... i.e. no error
          itemKeysDeleted.push(itemResult.key);
        }
      }
      if (itemKeysDeleted.length > 0) {
        //We found successfully deleted items, so remove them from the view.
        this._processDeletedItems(itemKeysDeleted);
      }
    }
    resultCallback(result); //Callback so that the caller can handle error UX.
  }

  /**
   * Called to update the item store when an item(s) has been updated.
   * @param itemKey : The item that was updated
   * @param updatedFields : The value of the fields tha were updated in the item
   * @param valueBeforeSave : optional: If this was called from updateListItemField, then the single field value before the API call was made
   * @param field : optional: If this was called from updateListItemField, then the IField schema for that field.
   */
  private _notifyItemStoreOfFieldUpdates(params: INotifyStoreOfUpdatesParams[]) {
    const items: ISPListRowWithErrors[] = [];
    const itemStatusesToUpdate: IListItemStatusUpdateParams<string, IListItemStatus>[] = [];
    const itemKeysToDelete: string[] = [];

    for (const param of params) {
      const { itemKey, updatedFields, valueBeforeSave, field, listRow } = param;
      let item = this._listItemStore.getItem(itemKey);
      let fieldsWithErrors: ListItemMessageMap =
        { ...this._listItemStore.getItemStatus(itemKey)?.fieldsWithErrors } || {};

      if (!item) {
        return;
      }

      let hasError = false;
      for (let i = 0; i < updatedFields.length; i++) {
        hasError =
          this._updateSingleFieldData(item, updatedFields[i], fieldsWithErrors, valueBeforeSave, field) ||
          hasError;
      }
      /**
       * Assign `listRow` only if there are no errors. If any field has an error, we want the error to be surfaced to the user.
       * We only want to assign `listRow` if there are no errors because if a field has an error, `listRow` will contain
       * previous values from before the item update.
       */
      item = {
        ...item,
        ...(!hasError ? listRow : {})
      };

      if (hasError) {
        itemStatusesToUpdate.push({
          itemKey,
          status: { hasError: true, fieldsWithErrors: fieldsWithErrors }
        });
      } else {
        itemKeysToDelete.push(itemKey);
      }

      items.push(item);
    }

    this._listItemStore.updateItemsStatus('ListItemProvider', itemStatusesToUpdate);
    this._listItemStore.deleteItemsStatus(itemKeysToDelete);
    this._listItemStore.updateItems('ListItemProvider', items);
  }

  /**
   * Updates a single field for a single item in the item store. Called in a for-loop from _notifyItemStoreOfFieldUpdates
   * @param item
   * @param updatedField
   * @param valueBeforeSave
   * @param field
   */
  private _updateSingleFieldData(
    item: ISPListRowWithErrors,
    updatedField: IListItemFormUpdateValue,
    fieldsWithErrors: ListItemMessageMap,
    valueBeforeSave?: any,
    field?: IField
  ): boolean {
    let curFieldRealFieldName = updatedField.FieldName;
    let hasError: boolean = false;

    // All edited fields from updateListItemField should be updated in itemStore so changes
    // are reflected to the user
    if (valueBeforeSave && field && field.realFieldName === curFieldRealFieldName) {
      // If this function is called from updateListItemField, then the valueBeforeSave provided here is the value
      // of the properties set before the XHR was made to validateUpdateListItem. The return value from
      // validateUpdateListItem, does not always return the expected values that you need to update
      // the item store appropriately. So use valueBeforeSave as the source of truth, since we know that the XHR succeeded.
      // Also, some SP fields have multiple values, so update all of them
      let rawValueRealFieldName = curFieldRealFieldName + '.';
      if (field.type === ColumnFieldType.Boolean) {
        //The internal value for boolean fields is stored as "internalFieldName.value" as opposed to "internalFieldName."
        rawValueRealFieldName += 'value';
      }
      item[curFieldRealFieldName] = valueBeforeSave.value;
      if (valueBeforeSave.rawValue) {
        item[rawValueRealFieldName] = valueBeforeSave.rawValue;
      }
    }

    if (!updatedField.HasException) {
      // If the field succeeded to update, remove from errors
      if (fieldsWithErrors && fieldsWithErrors[updatedField.FieldName]) {
        delete fieldsWithErrors[updatedField.FieldName];
      }
    } else {
      hasError = true;
      fieldsWithErrors[updatedField.FieldName] = updatedField.ErrorMessage;
    }

    return hasError;
  }

  private _updateItemStore(item: ISPListRow, targetField: IField, newValue: any) {
    const itemKey = this._listItemStore.getItemKey(item);

    const fieldsWithErrors: ListItemMessageMap =
      { ...this._listItemStore.getItemStatus(itemKey)?.fieldsWithErrors } || {};

    const realFieldName = targetField.realFieldName;

    // Editing a field with error removes it from itemStatus store regardless
    // of whether or not it was a proper fix, it then gets validated
    // after the last field with error is edited
    if (fieldsWithErrors[realFieldName]) {
      delete fieldsWithErrors[realFieldName];
    }

    if (newValue.value) {
      item[realFieldName] = newValue.value;
      item[`${realFieldName}.`] = newValue.rawValue || newValue.value;
    } else {
      item[realFieldName] = newValue;
      item[`${realFieldName}.`] = newValue;
    }

    const oldStatus = this._listItemStore?.getItemStatus(itemKey);
    this._listItemStore.updateItemsStatus('ListItemProvider.updateListItem.updateItemsStatus', [
      { itemKey, status: { ...oldStatus, hasError: true, fieldsWithErrors } }
    ]);
    this._listItemStore.updateItems('ListItemProvider.updateListItem.updateItems', [item]);
  }
}

// Synchronous key for ListItemProvider
export const listItemProviderKey: ResourceKey<ListItemProvider> = createResolvedResourceKey(
  require('module').id,
  ListItemProvider,
  {
    pageContext: pageContextKey,
    listItemStore: listItemStoreKey,
    listItemSelectionStore: listItemSelectionStoreKey,
    listItemDataSource: listItemDataSourceKey,
    getItemKey: getItemKeyResourceKey
  }
);
