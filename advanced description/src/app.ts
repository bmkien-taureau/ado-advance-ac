import { workItemNotificationListenerService } from './services/workItemNotificationListener';
import * as SDK from "azure-devops-extension-sdk";
import {
    IWorkItemFieldChangedArgs,
    IWorkItemFormService,
    IWorkItemNotificationListener,
    WorkItemTrackingServiceIds,
  } from 'azure-devops-extension-api/WorkItemTracking/WorkItemTrackingServices';
import './css/advanced-ac.scss';
import { TinyMCE } from 'tinymce/tinymce';

declare const tinymce: TinyMCE;

SDK.init({
    applyTheme: true,
    loaded: false,
  })
  .then(async (): Promise<void> => {
    const workItemFormService = await SDK.getService<IWorkItemFormService>(
        WorkItemTrackingServiceIds.WorkItemFormService
      );
    const workItemNotificationListener = new workItemNotificationListenerService(workItemFormService);
      (window as any)['_updateService'] = workItemFormService;
    await SDK.register<IWorkItemNotificationListener>(SDK.getContributionId(), workItemNotificationListener);

    tinymce.init({
      selector:'textarea',
      height: '100%',
      //mode: "specific_textareas",
      entity_encoding: "raw",
      plugins: 'accordion emoticons fullscreen insertdatetime searchreplace quickbars table image link lists help code wordcount',
      toolbar: 'undo redo | blocks | bold italic | alignleft aligncentre alignright alignjustify | indent outdent | bullist numlist | link image | code',
      content_style: 'body { font-family:Helvetica,Arial,sans-serif; font-size:16px }',
      menubar: true,
      toolbar_location: 'top',
      fixed_toolbar_container: '#editor-toolbar',
      init_instance_callback: function (editor) {
          editor.on('Change', function (e) {
              var textarea = document.getElementById('advanced-ac');
              if(textarea != null){
                workItemFormService.setFieldValue('Microsoft.VSTS.Common.AcceptanceCriteria', editor.getContent());
              }
          });
          editor.on('focus', function (e) {
            workItemNotificationListener.IgnoreSync(true);
          });
          editor.on('blur', function (e) {
            workItemNotificationListener.IgnoreSync(false);
          });
        }
    });

    await SDK.notifyLoadSucceeded();
});
