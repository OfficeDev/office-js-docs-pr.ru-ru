---
title: Наборы обязательных элементов общего API для Office
description: Узнайте больше о наборах требований к общим API для Office.
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f9929cb2f3de6499145540e12d1d96c55b24b1aa
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293522"
---
# <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке. Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).

> [!TIP]
> Ищете наборы требований API для *конкретных приложений* ? см. ниже.
>
> - [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md) (ExcelApi)
> - [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md) (WordApi)
> - [Наборы обязательных элементов API JavaScript для OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
> - [Наборы обязательных элементов PowerPoint JavaScript API](powerpoint-api-requirement-sets.md) (PowerPointApi)
> - [Общие сведения о наборах обязательных элементов API Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint. В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.

## <a name="common-api-requirement-sets"></a>Наборы обязательных элементов общего API

В следующих разделах перечислены общие наборы требований API, методы в каждом наборе и клиентские приложения Office, поддерживающие этот набор требований. Все эти наборы обязательных элементов API имеют версию 1.1, если не указано иное.

> [!TIP]
> Требуются сведения о том, где в приложениях и версиях Office поддерживаются надстройки и наборы требований? Сведения [о доступности клиентских приложений и платформ Office для надстроек Office](../../overview/office-add-in-availability.md).

### <a name="activeview"></a>ActiveView

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

См. статью [Наборы обязательных элементов для команд надстроек](add-in-commands-requirement-sets.md).

---

### <a name="bindingevents"></a>BindingEvents

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Binding.addHandlerAsync<br>Binding.removeHandlerAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel 2016 и более поздних версий в Windows<br>Excel в Интернете<br>Excel 2016 и более поздних версий на компьютерах Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает вывод в формате Office Open XML (OOXML) в виде байтового массива<br>(Office.FileType.Compressed) при использовании метода Document.getFileAsync.|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. статью [Наборы обязательных элементов API диалоговых окон](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>Файл

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| OneNote в Интернете<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Поддерживает приведение в HTML (Office.CoercionType.Html) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="identityapi"></a>IdentityAPI

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. статью [Наборы обязательных элементов API удостоверений](identity-api-requirement-sets.md). | Auth.getAccessToken |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. статью [Наборы требований к приведению изображений](image-coercion-requirement-sets.md). | Метод Document.setSelectedDataAsync|

---

### <a name="mailbox"></a>Mailbox

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
|Outlook для Windows<br>Outlook в Интернете<br>Outlook для Android<br>Outlook для Mac<br>Outlook для iOS|См. статью [Общие сведения о наборах обязательных элементов API для Outlook](outlook-api-requirement-sets.md).|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word для Windows<br>Word в Интернете<br>Word на iPad<br>Word для Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает приведение в структуру данных "матрица" (массив массивов, Office.CoercionType.Matrix) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает приведение в формат Open Office XML (OOXML, Office.CoercionType.Ooxml) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access||

---

### <a name="pdffile"></a>PdfFile

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает вывод в формате PDF (Office.FileType.Pdf)<br>при использовании метода Document.getFileAsync.|

---

### <a name="ribbonapi"></a>риббонапи

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См.: [наборы требований API ленты](ribbon-api-requirement-sets.md). | Office. Ribbon. Рекуеступдате |

---

### <a name="selection"></a>Selection

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Project для Windows<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Параметры

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="sharedruntime"></a>шаредрунтиме

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Ознакомьтесь с [общими наборами требований среды выполнения](shared-runtime-requirement-sets.md). | Office. AddIn. Жетстартупбехавиор<br>Office. AddIn. Hide<br>Office. AddIn. Онвисибилитимодечанжед<br>Office. AddIn. Сетстартупбехавиор<br>Office. AddIn. Шовастаскпане<br> |

---

### <a name="tablebindings"></a>TableBindings

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает приведение в структуру данных "таблица" (Office.CoercionType.Table) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="textbindings"></a>TextBindings

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии и Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Project для Windows<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает приведение в текстовый формат (Office.CoercionType.Text) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="textfile"></a>TextFile

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает вывод в текстовом формате (Office.FileType.Text) при использовании метода Document.getFileAsync.|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Методы, отсутствующие в наборе требований

Следующие методы в API JavaScript для Office не входят в набор требований. Если вашей надстройке необходимы какие-либо из этих методов, используйте элементы **Methods** и **Method** в манифесте надстройки, чтобы объявить их обязательными, или выполняйте проверку в среде выполнения с использованием оператора `if`. Дополнительную информацию можно узнать в статье [Указание приложений Office и требований к API](../../develop/specify-office-hosts-and-api-requirements.md).

|**Имя метода**|**Поддержка приложений Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Веб-приложения Access, Excel для Windows, Excel в Интернете, Excel на iPad и Excel для Mac|
|Document.getFilePropertiesAsync|Excel для Windows, Excel в Интернете, Excel на iPad, Excel для Mac, PowerPoint для Windows, PowerPoint в Интернете, PowerPoint на iPad, PowerPoint для Mac, Word для Windows, Word в Интернете, Word на iPad и Word для Mac|
|Document.getProjectFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getResourceFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedResourceAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedTaskAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedViewAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getTaskAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getTaskFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.goToByIdAsync|Excel для Windows, Excel в Интернете, Excel на iPad, Excel для Mac, PowerPoint для Windows, PowerPoint в Интернете, PowerPoint на iPad, PowerPoint для Mac, Word для Windows, Word в Интернете, Word на iPad и Word для Mac|
|Settings.addHandlerAsync|Веб-приложения Access и Excel в Интернете|
|Settings.refreshAsync|Веб-приложения Access, Excel для Windows, Excel в Интернете, PowerPoint для Windows, PowerPoint в Интернете, Word и Word в Интернете|
|Settings.removeHandlerAsync|Веб-приложения Access и Excel в Интернете|
|TableBinding.clearFormatsAsync|Excel для Windows, Excel в Интернете, Excel на iPad и Excel для Mac|
|TableBinding.setFormatsAsync|Excel для Windows, Excel в Интернете, Excel на iPad и Excel для Mac|
|TableBinding.setTableOptionsAsync|Excel для Windows, Excel в Интернете, Excel на iPad и Excel для Mac|

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и требований к API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
