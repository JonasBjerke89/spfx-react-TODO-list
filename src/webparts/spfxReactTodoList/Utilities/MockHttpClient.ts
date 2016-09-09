import { ITask } from './../components/SpfxReactTodoList';

export default class MockHttpClient {
  public static _results: ITask[] = [];

  public static get(restURL: string, options?: any) : Promise<ITask[]> {
    return new Promise<ITask[]>((resolve) => {
      resolve(MockHttpClient._results);
    });
  }
}