---
title: Используйте диалоговое окно "Office" для воспроизведения видео
description: Узнайте, как открыть и сыграть видео в диалоговом Office диалоговом окне
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 2519b2f105503a0479eee07d885a1543f5455343
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349884"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>Чтобы показать видео, Office диалоговое окно

В этой статье рассказывается, как играть видео в диалоговом окне Office надстройки.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования диалогового окна Office, как описано в статье [Использование API](dialog-api-in-office-add-ins.md)диалоговых Office в Office надстройки .

Чтобы играть видео в диалоговом окне с API Office диалоговом окне, выполните следующие действия:

1. Создайте страницу, содержащую iframe и отсутствие другого контента. Страница должна быть в том же домене, что и хост-страница. Напоминая о том, что такое хост-страница, см. в странице Откройте диалоговое [окно с хост-страницы.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) В `src` атрибуте iframe указать URL-адрес онлайн-видео. URL-адрес видео должен быть защищен с помощью протокола HTTPS. В этой статье мы назовем эту страницу "video.dialogbox.html". Ниже приведен пример разметки.

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Используйте вызов `displayDialogAsync` на главной странице, чтобы открыть страницу video.dialogbox.html.
3. Если надстройке необходимо знать, когда пользователь закрывает диалоговое окно, зарегистрируйте обработок события и обработите событие `DialogEventReceived` 12006. Подробные сведения см. в материале [Errors and events in the Office диалоговом окне](dialog-handle-errors-events.md).

Пример видео, играемого в диалоговом окне, см. в примере шаблона дизайна [видео-placemat.](../design/first-run-experience-patterns.md#video-placemat)

![Снимок экрана, показывающий воспроизведение видео в диалоговом окне надстройки перед Excel.](../images/video-placemats-dialog-open.png)
