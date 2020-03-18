---
title: Наборы обязательных элементов общего API для Office
description: Дополнительные сведения о наборах требований для общих API Office
ms.date: 07/17/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 39358b26547a464b9bb1b96f571bac7741e1c32d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717469"
---
# <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Сведения о поддержке надстроек ведущим приложением Office см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../../overview/office-add-in-availability.md).

Наборы обязательных элементов API *для конкретных ведущих приложений* см. ниже.

- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md) (WordApi)
- [Наборы обязательных элементов API JavaScript для OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Наборы обязательных элементов PowerPoint JavaScript API](powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Общие сведения о наборах обязательных элементов API Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint. В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.

## <a name="common-api-requirement-sets"></a>Наборы обязательных элементов общего API

В приведенных ниже разделах приводится список наборов обязательных элементов общего API, ведущие приложения Office, которые их поддерживают, и методы в каждом наборе. Все эти наборы обязательных элементов API имеют версию 1.1, если не указано иное.

### <a name="activeview"></a>ActiveView

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

См. статью [Наборы обязательных элементов для команд надстроек](add-in-commands-requirement-sets.md).

---

### <a name="bindingevents"></a>BindingEvents

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Binding.addHandlerAsync<br>Binding.removeHandlerAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Поддерживает вывод в формате Office Open XML (OOXML) в виде байтового массива<br>(Office.FileType.Compressed) при использовании метода Document.getFileAsync.|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. статью [Наборы обязательных элементов API диалоговых окон](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>Файл

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| OneNote в Интернете<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Поддерживает приведение в HTML (Office.CoercionType.Html) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="identityapi"></a>IdentityAPI

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. статью [Наборы обязательных элементов API удостоверений](identity-api-requirement-sets.md). | Auth.getAccessToken |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. статью [Наборы требований к приведению изображений](image-coercion-requirement-sets.md). | Метод Document.setSelectedDataAsync|

---

### <a name="mailbox"></a>Mailbox

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
|Outlook для Windows<br>Outlook в Интернете<br>Outlook для Android<br>Outlook для Mac<br>Outlook для iOS|См. статью [Общие сведения о наборах обязательных элементов API для Outlook](outlook-api-requirement-sets.md).|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word для Windows<br>Word в Интернете<br>Word для iPad<br>Word для Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Поддерживает приведение в структуру данных "матрица" (массив массивов, Office.CoercionType.Matrix) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает приведение в формат Open Office XML (OOXML, Office.CoercionType.Ooxml) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access||

---

### <a name="pdffile"></a>PdfFile

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Поддерживает вывод в формате PDF (Office.FileType.Pdf)<br>при использовании метода Document.getFileAsync.|

---

### <a name="selection"></a>Selection

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Project для Windows<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Параметры

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="sharedruntime"></a>шаредрунтиме

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Ознакомьтесь с [общими наборами требований среды выполнения](shared-runtime-requirement-sets.md). | Office. AddIn. Жетстартупбехавиор<br>Office. AddIn. Hide<br>Office. AddIn. Онвисибилитимодечанжед<br>Office. AddIn. Сетстартупбехавиор<br>Office. AddIn. Шовастаскпане<br> |

---

### <a name="tablebindings"></a>TableBindings

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает приведение в структуру данных "таблица" (Office.CoercionType.Table) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="textbindings"></a>TextBindings

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>Word 2013 и более поздней версии и Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Project для Windows<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word для iPad|Поддерживает приведение в текстовый формат (Office.CoercionType.Text) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|

---

### <a name="textfile"></a>TextFile

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает вывод в текстовом формате (Office.FileType.Text) при использовании метода Document.getFileAsync.|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Методы, отсутствующие в наборе требований

Следующие методы в API JavaScript для Office не входят в набор требований. Если вашей надстройке необходимы какие-либо из этих методов, используйте элементы **Methods** и **Method** в манифесте надстройки, чтобы объявить их обязательными, или выполняйте проверку в среде выполнения с использованием оператора `if`. Дополнительные сведения см. в статье [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md).

|**Имя метода**|**Поддержка ведущих приложений Office**|
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
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
