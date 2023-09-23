import {
    IWorkItemFieldChangedArgs,
    IWorkItemFormService,
    IWorkItemNotificationListener,
    WorkItemTrackingServiceIds,
    IWorkItemLoadedArgs
  } from 'azure-devops-extension-api/WorkItemTracking/WorkItemTrackingServices';
  import { TinyMCE } from '../tinymce/tinymce';

  declare const tinymce: TinyMCE;

export class workItemNotificationListenerService implements IWorkItemNotificationListener {
  private _workItemService: IWorkItemFormService;
  private _ignoreSync: boolean = false;
  public constructor(workItemService: IWorkItemFormService) {
        this._workItemService = workItemService;
    }

    public  onLoaded = async (loadedArgs: IWorkItemLoadedArgs) => {
      console.log('IWorkItemNotificationListener.onLoaded - called.', loadedArgs);
      this.LoadDescription();
    };

    public  onSaved = async () => console.log('IWorkItemNotificationListener.onSaved - called.');
    public  onRefreshed = async () => console.log('IWorkItemNotificationListener.onRefreshed - called.');
    public  onReset = async () => console.log('IWorkItemNotificationListener.onReset - called.');
    public  onUnloaded = async () => console.log('IWorkItemNotificationListener.onUnloaded - called.');

    public  onFieldChanged = async (fieldChangedArgs: IWorkItemFieldChangedArgs) => {
      //console.log('IWorkItemNotificationListener.onFieldChanged - called.', fieldChangedArgs);
      if(!this._ignoreSync)
        tinymce.get('advanced-description')?.setContent(fieldChangedArgs.changedFields['System.Description']);
    };

    public IgnoreSync(ignore: boolean) {
       this._ignoreSync = ignore;
    }

    LoadDescription() {
      this._workItemService.getFieldValue('System.Description', {
        returnOriginalValue: false
      }).then((val)=>{
          var content = val.toString();
          var textarea = document.getElementById('advanced-description');
          if(textarea != null) {
            textarea.innerHTML = content;
          }
      });
    }
}