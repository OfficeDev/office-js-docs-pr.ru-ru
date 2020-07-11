---
title: Версии Office и наборы обязательных элементов
description: Поддерживаемые платформы Office.js с использованием JavaScript API
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 02f3d91256ea05e526ebe2e4e4090b1908d7292a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093583"
---
# <a name="office-versions-and-requirement-sets"></a>Версии Office и наборы обязательных элементов

There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in. 

> [!NOTE]
> - Office работает на различных платформах, в том числе Windows, в браузере, на компьютерах Mac и на iPad.
> - Примеры ведущих приложений Office — Excel, Word, PowerPoint, Outlook, OneNote и другие продукты.  
> - Набор обязательных элементов — это именованная группа элементов API, например, `ExcelApi 1.5`, `WordApi 1.3` и т. д.  

## <a name="how-to-check-your-office-version"></a>Как узнать, какая версия Office используется

To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):

![Проверка версии Office](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Доступность наборов обязательных элементов для Office

Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).

Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

Кроме того, к общему API были добавлены другие функции, такие как команды надстроек (расширение ленты) и возможность запуска диалоговых окон (API диалоговых окон). Наборы обязательных элементов для команд надстроек и API диалоговых окон — это наборы API, общие для всех ведущих приложений Office.

An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:

- [Наборы обязательных элементов API JavaScript для Excel](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)
- [Наборы обязательных элементов API JavaScript для Word](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)
- [Наборы обязательных элементов API JavaScript для OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)
- [Наборы обязательных элементов PowerPoint JavaScript API](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Общие сведения о наборах обязательных элементов API Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md) (MailBox)

Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:

- [Общие наборы обязательных элементов для Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Наборы обязательных элементов для команд надстроек](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Наборы обязательных элементов API диалоговых окон](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Наборы обязательных элементов API идентификации](../reference/requirement-sets/identity-api-requirement-sets.md)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

Библиотека API JavaScript для Office (Office.js) включает все доступные наборы обязательных элементов. Наборы обязательных элементов `ExcelApi 1.3` и `WordApi 1.3` существуют, но набора обязательных элементов `Office.js 1.3` нет. Доступ к последней версии Office.js осуществляется через единую конечную точку Office, интегрированную в сеть доставки содержимого (CDN). Дополнительные сведения о CDN Office.js, в том числе об управлении версиями и обратной совместимости, см. в статье [Общие сведения об интерфейсе API JavaScript для Office](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-hosts-and-requirement-sets"></a>Указание ведущих приложений Office и наборов обязательных элементов

There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>См. также

- [Указание ведущих приложений Office и обязательных элементов API](../develop/specify-office-hosts-and-api-requirements.md)
- [Установка последней версии Office](../develop/install-latest-office-version.md)
- [Обзор каналов обновления для Приложений Microsoft 365](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Получите максимум от наших продуктов благодаря Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
