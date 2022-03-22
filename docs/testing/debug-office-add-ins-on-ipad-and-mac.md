---
title: Отладка надстроек Office на Mac
description: Узнайте, как использовать Mac для отлаговки Office надстроек.
ms.date: 03/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: dc9017eb7bd27ee0bc22d3ad448e5996895c5eee
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711211"
---
# <a name="debug-office-add-ins-on-a-mac"></a>Отладка надстроек Office на Mac

Так как надстройки создаются с помощью кода HTML и JavaScript, они рассчитаны на работу на многих платформах, но отрисовка HTML в разных браузерах может слегка отличаться. В этой статье описывается отладка надстроек на компьютере Mac.

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Отладка с помощью Safari Web Inspector на компьютере Mac

Если у вас есть надстройка, которая отображает пользовательский интерфейс в области задач или контентной надстройке, вы можете отлаживать надстройку Office с помощью Safari Web Inspector.

Чтобы иметь возможность отлагоравить Office на Mac, необходимо иметь Mac OS High Sierra и Mac Office версии 16.9.1 (сборка 18012504) или более поздней версии. Если у вас нет сборки Office Mac, вы можете получить ее, присоединившись к Microsoft 365 [разработчика](https://developer.microsoft.com/office/dev-program).

Для этого откройте терминал и установите свойство `OfficeWebAddinDeveloperExtras` для соответствующего приложения Office следующим образом:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > Сборки mac App Store Office не поддерживают флаг`OfficeWebAddinDeveloperExtras`.

Затем откройте приложение Office и [загрузите свою неопубликованную надстройку](sideload-an-office-add-in-on-ipad-and-mac.md). Щелкните надстройку правой кнопкой мыши. В контекстном меню отобразится пункт **Проверить элемент**. Выберите его. Он появится в инспекторе, где можно устанавливать точки останова и отлаживать надстройку.

> [!NOTE]
> Если при попытке использовать инспектор диалоговое окно мерцает, обновите Office до последней версии. Если это не устраняет мерцание, попробуйте следующее обходное решение.
>
> 1. Уменьшите размер диалогового окна.
> 1. Выберите пункт **Проверить элемент**, который откроется в новом окне.
> 1. Измените размер диалогового окна на исходный.
> 1. Используйте инспектор должным образом.

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>Очистка кэша приложения Office на компьютере Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
