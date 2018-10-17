# <a name="office-common-api-requirement-sets"></a>Стандартные наборы требований API для Office

Наборы требований — это именованные группы элементов API. С помощью наборов требований, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Нужны сведения о поддержке надстроек ведущим приложением Office? См. статью [Доступность ведущих приложений и платформ для надстроек Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

Ищите наборы требований API *для конкретных ведущих приложений*? См. следующие наборы требований API:
 
- [Наборы требований API JavaScript для Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Наборы требований API JavaScript для Word](word-api-requirement-sets.md) (WordApi)
- [Наборы требований API JavaScript для OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Общие сведения о наборах требований API для Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint. В качестве альтернативы мы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.

## <a name="common-api-requirement-sets"></a>Стандартные наборы требований API

В приведенной ниже таблице указаны стандартные наборы требований API, ведущие приложения Office, которые их поддерживают, и методы в каждом наборе. Все эти наборы требований API имеют версию 1.1.

|**Набор требований**|**Ведущее приложение Office**|**Методы в наборе**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac|Document.getActiveViewAsync|
| AddInCommands | См. статью [Наборы требований для команд надстроек](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Веб-приложения Access<br>Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает вывод в формате Office Open XML (OOXML) в виде байтового массива<br>(Office.FileType.Compressed) при использовании метода Document.getFileAsync.|
| CustomXmlParts    | Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | См. статью [Наборы требований API Dialog](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в формат HTML (Office.CoercionType.Html) при чтении и записи данных с использованием методов Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync или методы Binding.setDataAsync.|
| IdentityAPI | См. статью [Наборы требований API Identity](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel для iPad<br>Excel для Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает преобразование в изображение (Office.CoercionType.Image) при записи данных с помощью метода Document.setSelectedDataAsync.|
| Почтовый ящик   |Outlook для Windows<br>Outlook для веб-браузеров<br>Outlook для Android<br>Outlook для Mac<br>Веб-приложение Outlook |См. статью [Общие сведения о наборах требований API для Outlook](outlook-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>Word<br>Word Online<br>Word для iPad<br>Word для Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в структуру «матрица» (массив массивов, Office.CoercionType.Matrix) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync или Binding.setDataAsync.|
| OoxmlCoercion | Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в формат Open Office XML (OOXML, Office.CoercionType.Ooxml) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync или Binding.setDataAsync.|
| PartialTableBindings  | Веб-приложения Access||
| PdfFile   | Excel для Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает вывод в формате PDF (Office.FileType.Pdf)<br>при использовании метода Document.getFileAsync.|
| Selection | Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Project<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Веб-приложения Access<br>Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Веб-приложения Access<br>Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Веб-приложения Access<br>Excel<br>Excel Online<br>Excel для iPad<br>Excel для Mac<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в структуру данных "таблица" (Office.CoercionType.Table) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync или Binding.setDataAsync.|
| TextBindings  | Excel<br>Excel Online<br>Excel для iPad<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel для iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint для iPad<br>PowerPoint для Mac<br>Project<br>Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в текстовый формат (Office.CoercionType.Text) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync или Binding.setDataAsync.|
| TextFile  | Word 2013 и более поздних версий<br>Word 2016 и более поздних версий для Mac<br>Word Online<br>Word для iPad|Поддерживает вывод в текстовом формате (Office.FileType.Text) при использовании метода Document.getFileAsync.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Методы, отсутствующие в наборе требований

Следующие методы API JavaScript для Office не входят в набор требований. Если вашей надстройке необходимы какие-либо из этих методов, используйте элементы **Methods** и **Method** в манифесте надстройки, чтобы объявить их обязательными, или выполните проверку в среде выполнения с использованием оператора  `if`. Дополнительные сведения см. в статье [Указание ведущих приложений Office и требований API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Имя метода**|**Поддержка ведущих приложений Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Веб-приложениях Access, Excel, Excel Online и Excel для iPad|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel для iPad, Excel для Mac, PowerPoint, PowerPoint Online, PowerPoint для iPad, PowerPoint для Mac, Word, Word Online, Word для iPad и Word для Mac|
|Document.getProjectFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getResourceFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedResourceAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedTaskAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedViewAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getTaskAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getTaskFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel для iPad, Excel для Mac, PowerPoint, PowerPoint Online, PowerPoint для iPad, PowerPoint для Mac, Word, Word Online, Word для iPad и Word для Mac|
|Settings.addHandlerAsync|Веб-приложениях Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word и Word Online|
|Settings.refreshAsync|Веб-приложениях Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word и Word Online|
|Settings.removeHandlerAsync|Веб-приложениях Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word и Word Online|
|TableBinding.clearFormatsAsync|Excel, Excel Online и Excel для Mac|
|TableBinding.setFormatsAsync|Excel, Excel Online, Excel для iPad и Excel для Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, Excel для iPad и Excel для Mac|

## <a name="see-also"></a>См. также

- [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и требований API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
