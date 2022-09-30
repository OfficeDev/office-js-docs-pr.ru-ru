---
title: Версии Office и наборы обязательных элементов
description: Поддерживаемые платформы Office.js с использованием JavaScript API
ms.date: 09/14/2022
ms.localizationpriority: high
ms.openlocfilehash: 669977f87974a1ec5519ddbbe3d38c5a290ec84f
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234909"
---
# <a name="office-versions-and-requirement-sets"></a>Версии Office и наборы обязательных элементов

Не все версии Office поддерживают все API в API JavaScript для Office (Office.js). Office 2013 в Windows — это последняя версия Office, которая поддерживала надстройки Office. Вы не всегда можете контролировать версию Office, установленную пользователями. Чтобы справиться с этой ситуацией, мы предоставляем систему, называемую наборами обязательных элементов, чтобы определить, поддерживает ли приложение Office возможности, необходимые в надстройке Office.

> [!NOTE]
>
> - Office работает на различных платформах, в том числе Windows, в браузере, на компьютерах Mac и на iPad.
> - Примерами приложений Office являются продукты Office: Excel, Word, PowerPoint, Outlook, OneNote и т. д.
> - Office доступен по подписке На Microsoft 365 или бессрочной лицензии. Бессрочная версия доступна в рамках соглашения корпоративного лицензирования или розничной торговли.
> - Набор обязательных элементов — это именованная группа членов API, например, `ExcelApi 1.5`и `WordApi 1.3`т. д.

## <a name="how-to-check-your-office-version"></a>Как узнать, какая версия Office используется

Чтобы определить используемую версию Office, в приложении Office откройте меню **Файл** и выберите **Учетная запись**. Версия Office отображается в разделе " **Сведения о продукте** ". Например, на следующем снимке экрана показана версия Office 1802 (сборка 9026.1000).

![Проверка версии Office.](../images/office-version.png)

> [!NOTE]
> Если ваша версия Office отличается от этой, см. статью "Какая у меня версия [Outlook?"](https://support.microsoft.com/office/b3a9568c-edb5-42b9-9825-d48d82b2257c) или "Сведения о [Office: какая версия Office я могу использовать?"](https://support.microsoft.com/topic/932788b8-a3ce-44bf-bb09-e334518b8b19) , чтобы понять, как получить эти сведения для вашей версии.

## <a name="office-requirement-sets-availability"></a>Доступность наборов обязательных элементов для Office

Надстройки Office могут использовать наборы обязательных элементов API, чтобы определить, поддерживает ли приложение Office необходимые элементы API. Поддержка набора обязательных элементов зависит от приложения Office и версии приложения Office (см. предыдущий раздел ["Проверка версии Office"](#how-to-check-your-office-version)).

Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

Кроме того, к общему API были добавлены другие функции, такие как команды надстроек (расширение ленты) и возможность запуска диалоговых окон (API диалоговых окон). Команды надстроек и наборы обязательных элементов API диалоговых окон — это примеры наборов API, совместно используемых различными приложениями Office.

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Наборы обязательных элементов API JavaScript для Excel](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [Наборы обязательных элементов API JavaScript для OneNote](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [Наборы обязательных элементов API JavaScript для Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (почтовый ящик)
- [Наборы обязательных элементов PowerPoint JavaScript API](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Наборы обязательных элементов API JavaScript для Word](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)

Некоторые наборы обязательных элементов содержат API, которые могут использоваться несколькими приложениями Office. Дополнительные сведения об этих наборах обязательных элементов см. в следующих статьях.

- [Общие наборы обязательных элементов для Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Наборы обязательных элементов для команд надстроек](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [Наборы обязательных элементов API диалоговых окон](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [Наборы обязательных элементов источников диалоговых окон](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)
- [Наборы обязательных элементов API идентификации](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Наборы обязательных элементов для приведения изображений](/javascript/api/requirement-sets/common/image-coercion-requirement-sets)
- [Наборы обязательных элементов cочетаний клавиш](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [Открытие набора обязательных элементов окна браузера](/javascript/api/requirement-sets/common/open-browser-window-api-requirement-sets)
- [Наборы обязательных элементов API ленты](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [Наборы обязательных элементов в среде выполнения](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>Указание приложений Office и наборов обязательных элементов

There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>См. также

- [Указание приложений Office и обязательных элементов API](../develop/specify-office-hosts-and-api-requirements.md)
- [Установка последней версии Office](../develop/install-latest-office-version.md)
- [Обзор каналов обновления для Приложений Microsoft 365](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Свежий взгляд на продуктивность благодаря Microsoft 365 и Microsoft Teams](https://products.office.com/compare-all-microsoft-office-products?tab=2)
