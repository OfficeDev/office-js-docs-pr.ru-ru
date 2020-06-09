---
title: Используйте диалоговое окно "Office" для воспроизведения видео
description: Сведения о том, как открыть и прослушать видео в диалоговом окне Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: e150206b60fdff852621971fd4417ff9bdfe7eb3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608169"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>Отображение видео с помощью диалогового окна Office

В этой статье объясняется, как воссоздать видео в диалоговом окне надстройки Office.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования диалогового окна Office, как описано в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md)надстройках Office.

Для проигрывания видео в диалоговом окне с помощью API диалогового окна Office выполните следующие действия:

1. Создание страницы, содержащей IFRAME, без другого контента. Страница должна находиться в том же домене, что и Главная страница. Напоминание о странице ведущего приложения можно узнать в разделе [Открытие диалогового окна на странице узла](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page). В `src` атрибуте IFRAME укажите URL-адрес видео в Интернете. URL-адрес видео должен быть защищен с помощью протокола HTTPS. В этой статье мы вызываем эту страницу "Video. DialogBox. HTML". Ниже приведен пример разметки:

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Используйте вызов `displayDialogAsync` на главной странице, чтобы открыть страницу video.dialogbox.html.
3. Если надстройка должна знать, когда пользователь закрывает диалоговое окно, зарегистрируйте обработчик для `DialogEventReceived` события и обработайте событие 12006. Дополнительные сведения: ["ошибки и события" в диалоговом окне Office](dialog-handle-errors-events.md).

Пример видеоконференций, воспроизводимого в диалоговом окне, приведен в статье [Образец оформления видео представление](../design/first-run-experience-patterns.md#video-placemat).

![Снимок экрана: диалоговое окно воспроизведения видео в диалоговом окне надстройки](../images/video-placemats-dialog-open.png)
