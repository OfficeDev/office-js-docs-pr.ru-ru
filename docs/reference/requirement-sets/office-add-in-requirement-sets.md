---
title: Наборы обязательных элементов общего API для Office
description: Дополнительные дополнительные Office общих наборов API.
ms.date: 04/28/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 959f03bf41496c1506087c2851efad336cdec676
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253349"
---
# <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

> [!TIP]
> Ищете *наборы требований* для API, определенных для приложений? см. ниже.
>
> - [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md) (ExcelApi)
> - [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md) (WordApi)
> - [Наборы обязательных элементов API JavaScript для OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
> - [Наборы обязательных элементов PowerPoint JavaScript API](powerpoint-api-requirement-sets.md) (PowerPointApi)
> - [Общие сведения о наборах обязательных элементов API Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint. В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.

## <a name="common-api-requirement-sets"></a>Наборы обязательных элементов общего API

В следующих разделах перечисляются общие наборы требований API, методы в каждом наборе и Office клиентские приложения, которые поддерживают этот набор требований. Все эти наборы обязательных элементов API имеют версию 1.1, если не указано иное.

> [!TIP]
> Нужна информация о том, где надстройки и наборы требований поддерживаются Office и версией? См. Office клиентского приложения и доступности [платформы для Office надстройки](../../overview/office-add-in-availability.md).

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
| Excel 2016 и более поздней Windows<br>Excel в Интернете<br>Excel 2016 и позднее на Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Поддерживает вывод в формате Office Open XML (OOXML) в виде байтового массива<br>(Office.FileType.Compressed) при использовании метода Document.getFileAsync.|

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

### <a name="file"></a>File

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

### <a name="openbrowserwindowapi"></a>OpenBrowserWindowApi

|**Ведущие приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. [наборы требований к API API окна открытого браузера.](open-browser-window-api-requirement-sets.md) | Office.context.ui.openBrowserWindow |

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access||

---

### <a name="pdffile"></a>PdfFile

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете|Поддерживает вывод в формате PDF (Office.FileType.Pdf)<br>при использовании метода Document.getFileAsync.|

---

### <a name="ribbonapi"></a>RibbonApi

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. [наборы требований к API ленты.](ribbon-api-requirement-sets.md) | Office.ribbon.requestUpdate |

---

### <a name="selection"></a>Selection

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Project для Windows<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Settings

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| Веб-приложения Access<br>Excel для Windows<br>Excel в Интернете<br>Excel на iPad<br>Excel для Mac<br>OneNote в Интернете<br>PowerPoint для Windows<br>PowerPoint в Интернете<br>PowerPoint на iPad<br>PowerPoint для Mac<br>Word 2013 и более поздней версии для Windows<br>Word 2016 и более поздней версии для Mac<br>Word в Интернете<br>Word на iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="sharedruntime"></a>SharedRuntime

|**Приложения Office**|**Методы в наборе**|
|:-----|:-----|
| См. [общие наборы требований к времени работы.](shared-runtime-requirement-sets.md) | Office.addin.getStartupBehavior<br>Office.addin.hide<br>Office.addin.onVisibilityModeChanged<br>Office.addin.setStartupBehavior<br>Office.addin.showAsTaskpane<br> |

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

Следующие методы в API Office JavaScript не являются частью набора требований. Если вашей надстройке необходимы какие-либо из этих методов, используйте элементы **Methods** и **Method** в манифесте надстройки, чтобы объявить их обязательными, или выполняйте проверку в среде выполнения с использованием оператора `if`. Дополнительные сведения см. в [Office приложениях и требованиях API.](../../develop/specify-office-hosts-and-api-requirements.md)

|**Имя метода**|**Office поддержки приложений**|
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
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
