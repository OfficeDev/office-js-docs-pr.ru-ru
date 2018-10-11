# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.

В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).

> [!IMPORTANT]
> Элемент SupportsSharedFolders доступен только в наборе требований предварительной версии надстроек Outlook относительно Exchange Online.
> Надстройки, использующие этот элемент, не разрешаются в магазине Office или для централизованного развертывания.

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Да      | Определяет, где предоставляется функциональность надстройки. |
| [FunctionFile](functionfile.md)       | Да      | URL-адрес файла, содержащего функции JavaScript.|
| [GetStarted](getstarted.md)           | Нет       | Определяет выноску, которая отображается при установке надстройки в основных приложениях Word, Excel и PowerPoint. |
| SupportsSharedFolders                 | Нет       | Определяет, доступна ли надстройка Outlook в сценарии делегата, и имеет значение *false* по умолчанию. Набор требований для предварительной версии.|

## <a name="desktopformfactor-example"></a>Пример DesktopFormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
