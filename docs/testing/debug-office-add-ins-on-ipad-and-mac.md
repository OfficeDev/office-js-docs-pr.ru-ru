---
title: Отладка надстроек Office на Mac
description: Узнайте, как использовать Mac для отладки надстроек Office.
ms.date: 03/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32d896743932abc7cf8be6bd62a491fc93fe0d1b
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713002"
---
# <a name="debug-office-add-ins-on-a-mac"></a>Отладка надстроек Office на Mac

Так как надстройки создаются с помощью кода HTML и JavaScript, они рассчитаны на работу на многих платформах, но отрисовка HTML в разных браузерах может слегка отличаться. В этой статье описывается отладка надстроек на компьютере Mac.

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Отладка с помощью Safari Web Inspector на компьютере Mac

Если у вас есть надстройка, которая отображает пользовательский интерфейс в области задач или контентной надстройке, вы можете отлаживать надстройку Office с помощью Safari Web Inspector.

Чтобы иметь возможность отладки надстроек Office на Mac, необходимо иметь Mac OS High Sierra и Mac Office версии 16.9.1 (сборка 18012504) или более поздней. Если у вас нет сборки Office mac, ее можно получить, присоединившись к [программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program).

Для этого откройте терминал и установите свойство `OfficeWebAddinDeveloperExtras` для соответствующего приложения Office следующим образом:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > Сборки App Store Mac не поддерживают этот флаг`OfficeWebAddinDeveloperExtras`.

Затем откройте приложение Office и [загрузите свою неопубликованную надстройку](sideload-an-office-add-in-on-mac.md). Щелкните надстройку правой кнопкой мыши. В контекстном меню отобразится пункт **Проверить элемент**. Выберите его. Он появится в инспекторе, где можно устанавливать точки останова и отлаживать надстройку.

> [!NOTE]
> Если при попытке использовать инспектор диалоговое окно мерцает, обновите Office до последней версии. Если это не устраняет мерцание, попробуйте выполнить следующее решение.
>
> 1. Уменьшите размер диалогового окна.
> 1. Выберите пункт **Проверить элемент**, который откроется в новом окне.
> 1. Измените размер диалогового окна на исходный.
> 1. Используйте инспектор должным образом.

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>Очистка кэша приложения Office на компьютере Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
