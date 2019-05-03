---
title: Отладка надстроек Office на Mac
description: ''
ms.date: 04/24/2019
localization_priority: Priority
ms.openlocfilehash: 6d77dd0d90e68c2147ffea67d12026fc194fa642
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517095"
---
# <a name="debug-office-add-ins-on-a-mac"></a>Отладка надстроек Office на Mac

Visual Studio подходит для разработки и отладки надстроек в Windows, но с его помощью невозможно выполнять отладку надстроек на компьютере Mac. Так как надстройки создаются с помощью кода HTML и JavaScript, они рассчитаны на работу на многих платформах, но отрисовка HTML в разных браузерах может слегка отличаться. В этой статье описывается отладка надстроек на компьютере Mac.

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Отладка с помощью Safari Web Inspector на компьютере Mac

Если у вас есть надстройка, которая отображает пользовательский интерфейс в области задач или контентной надстройке, вы можете отлаживать надстройку Office с помощью Safari Web Inspector.

Отладку надстроек Office на компьютере Mac можно выполнить, только если на нем установлена система Mac OS High Sierra И Office для Mac версии 16.9.1 (сборка 18012504) или более поздней. Если у вас нет сборки Office для Mac, вы можете получить ее, присоединившись к [программе для разработчиков Office 365](https://aka.ms/o365devprogram).

Для этого откройте терминал и установите свойство `OfficeWebAddinDeveloperExtras` для соответствующего приложения Office следующим образом:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Затем откройте приложение Office и [загрузите свою неопубликованную надстройку](sideload-an-office-add-in-on-ipad-and-mac.md). Щелкните надстройку правой кнопкой мыши. В контекстном меню отобразится пункт **Проверить элемент**.  Выберите его. Он появится в инспекторе, где можно устанавливать точки останова и отлаживать надстройку.

> [!NOTE]
> Если при попытке использовать инспектор диалоговое окно мерцает, обновите Office до последней версии. Если проблема с мерцанием сохраняется, попробуйте применить следующее временное решение:
> 1. Уменьшите размер диалогового окна.
> 2. Выберите пункт **Проверить элемент**, который откроется в новом окне.
> 3. Измените размер диалогового окна на исходный.
> 4. Используйте инспектор должным образом.

## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Очистка кэша приложения Office на компьютере Mac или iPad

Для повышения производительности надстройки часто кэшируются в Office для Mac. Как правило, для очистки кэша необходимо перезагрузить надстройку. Если в одном документе несколько надстроек, автоматическая очистка кэша может не сработать при перезагрузке.

На компьютере Mac можно очистить кэш вручную, удалив все содержимое папки `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

На iPad можно вызвать в надстройке метод JavaScript `window.location.reload(true)` для принудительной перезагрузки. Вы также можете переустановить Office.
