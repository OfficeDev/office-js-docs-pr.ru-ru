# <a name="supportssharedfolders-element"></a>Элемент SupportsSharedFolders

Он определяет, является ли надстройка Outlook доступной в сценарии делегирования. Элемент **SupportsSharedFolders** является дочерним элементом элемента [DesktopFormFactor](desktopformfactor.md). Он имеет значение *false* по умолчанию.

> [!IMPORTANT]
> Этот элемент доступен только в [Наборе требований предварительного просмотра надстроек Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) по отношению к Exchange Online. Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.

Ниже приведен пример использования элемента **SupportsSharedFolders** .

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
