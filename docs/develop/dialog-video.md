---
title: Используйте диалоговое окно "Office" для воспроизведения видео
description: Сведения о том, как открыть и прослушать видео в диалоговом окне Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: f0d524996b105061b8e5d1b584d8b3e0d44eec7c
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131775"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>Отображение видео с помощью диалогового окна Office

В этой статье объясняется, как воссоздать видео в диалоговом окне надстройки Office.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования диалогового окна Office, как описано в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md)надстройках Office.

Для проигрывания видео в диалоговом окне с помощью API диалогового окна Office выполните следующие действия:

1. Создание страницы, содержащей IFRAME, без другого контента. Страница должна находиться в том же домене, что и Главная страница. Напоминание о странице ведущего приложения можно узнать в разделе [Открытие диалогового окна на странице узла](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page). В `src` атрибуте IFRAME укажите URL-адрес видео в Интернете. URL-адрес видео должен быть защищен с помощью протокола HTTPS. В этой статье мы вызываем эту страницу "video.dialogbox.html". Ниже приведен пример разметки:

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Используйте вызов `displayDialogAsync` на главной странице, чтобы открыть страницу video.dialogbox.html.
3. Если надстройка должна знать, когда пользователь закрывает диалоговое окно, зарегистрируйте обработчик для `DialogEventReceived` события и обработайте событие 12006. Дополнительные сведения: ["ошибки и события" в диалоговом окне Office](dialog-handle-errors-events.md).

Пример видеоконференций, воспроизводимого в диалоговом окне, приведен в статье [Образец оформления видео представление](../design/first-run-experience-patterns.md#video-placemat).

![Снимок экрана: воспроизведение видео в диалоговом окне надстройки перед Excel](../images/video-placemats-dialog-open.png)
