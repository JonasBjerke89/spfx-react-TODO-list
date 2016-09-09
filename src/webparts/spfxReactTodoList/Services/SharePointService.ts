import { ITODOList } from './../components/SpfxReactTodoList';
import SPRestAPIClient from './../Utilities/SPRestAPIClient';

export default class SharePointService
{
  public static getTodos(apiUrl:string, listName: string, hideFinishedTasks: boolean): Promise<ITODOList> {

    return new Promise<ITODOList>((resolve: (todos: ITODOList) => void, reject: (err: any) => void): void => {
      var url : string = apiUrl + '/_api/web/lists/getbytitle(\'' + listName + '\')/items?$select=Id,Title,Status,Author/Title,Author/SipAddress,Created&$orderby=ID desc&$expand=Author';

      if (hideFinishedTasks === true) {
        url += "&$filter=Status ne 'Completed'";
      }

      SPRestAPIClient.Request(url).then((response : any) => {

        const todos: ITODOList = {
            value:
            [

            ]
            };

            for (let i: number = 0; i < response.value.length; i++) {
              todos.value.push({ title: response.value[i].Title, completed:false});
            }

        resolve(todos);
      }, (error: any) => {
        reject(error);
      });
    });
  }

  public static addTodo(apiUrl:string, listName: string, title: string): Promise<{}> {

    return new Promise<{}>((resolve: () => void, reject: (err: any) => void): void => {
      this.getRequestDigest(apiUrl).then((digest) => {
          var url : string = apiUrl + '/_api/web/lists/getbytitle(\'' + listName + '\')/items?$select=Id,Title,Status,Author/Title,Author/SipAddress,Created&$orderby=ID desc&$expand=Author';

          var headers = {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': digest
          }

          var body: string = JSON.stringify({
            '__metadata': { 'type': 'SP.Data.' + listName.charAt(0).toUpperCase() + listName.slice(1) + 'ListItem' },
            'Title': title
          });

          SPRestAPIClient.Request(url, 'POST', headers, body).then((response : any) => {
            resolve();
          }, (error: any) => {
            reject(error);
          });
      });
    });
  }

  private static getRequestDigest(siteUrl: string): Promise<string> {

    return new Promise<string>((resolve, reject): void => {
      SPRestAPIClient.Request(`${siteUrl}/_api/contextinfo`, 'POST').then((data: { FormDigestValue: string }): void => {
        resolve(data.FormDigestValue);
      }, (error: any): void => {
        reject(error);
      });
    });
  }
}




/*public addTodo(todo: string, sharePointApi: string, todoListName: string): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    this.$http({
      url: sharePointApi + 'contextinfo',
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((digestResult: ng.IHttpPromiseCallbackArg<{ FormDigestValue: string }>): void => {
      const requestDigest: string = digestResult.data.FormDigestValue;
      const body: string = JSON.stringify({
        '__metadata': { 'type': 'SP.Data.' + todoListName.charAt(0).toUpperCase() + todoListName.slice(1) + 'ListItem' },
        'Title': todo
      });
      this.$http({
        url: sharePointApi + 'web/lists/getbytitle(\'' + todoListName + '\')/items',
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest
        },
        data: body
      }).then((result: ng.IHttpPromiseCallbackArg<{}>): void => {
        deferred.resolve();
      });
    });

    return deferred.promise;
  }*/