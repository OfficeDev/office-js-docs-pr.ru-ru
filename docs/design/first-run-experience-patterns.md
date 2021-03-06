---
title: Шаблоны интерфейса первого запуска для надстроек Office
description: Узнайте о лучших практиках разработки первого запуска в Office надстройки.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: d020a281aca10805ba8fd1176403f3788f6d716c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076345"
---
# <a name="first-run-experience-patterns"></a>Шаблоны интерфейса первого запуска

Интерфейс первого запуска (FRE) обеспечивает знакомство пользователя с вашей надстройкой. Когда пользователь впервые открывает надстройку, отображается интерфейс FRE, который дает им представление о функциях, возможностях и/или преимуществах надстройки. Этот интерфейс формирует первое впечатление от надстройки и может сильно повлиять на вероятность того, что пользователь вернется и продолжит пользоваться вашей надстройкой.

## <a name="best-practices"></a>Рекомендации

Следуйте этим рекомендациям при создании интерфейса первого запуска:

|Правильно|Неправильно|
|:------|:------|
|Ясно и кратко опишите основные действия в надстройке. | Не указывайте сведения, не относящиеся к началу работы.
|Предоставьте пользователям возможность выполнить действие, которое создаст у них положительное впечатление от использования надстройки. | Не следует ожидать, что пользователи изучат все возможности сразу. Сосредоточьтесь на самом ценном действии.
|Создайте привлекательный интерфейс, в котором пользователи захотят выполнить все действия. | Не заставляйте пользователей просматривать весь интерфейс первого запуска. Предоставьте пользователям возможность обойти его. |

Решите, как часто необходимо применять интерфейс, используемый при первом запуске: один раз или периодически. Например, если ваша надстройка используется только время от времени, пользователи могут забывать ее возможности, и тогда им будет полезно еще раз ознакомиться с интерфейсом первого запуска.

При создании или улучшении интерфейса первого запуска для надстройки применяйте указанные ниже шаблоны.

## <a name="carousel"></a>Карусель

Карусель знакомит пользователей с рядом функций или предоставляет определенные сведения, прежде чем они начнут использовать надстройку.

*Рис. 1. Разрешить пользователям заранее или пропустить начало страниц потока карусель*

![Иллюстрация, показывающая шаг 1 карусели в первом запуске области задач Office настольного приложения. В этом примере действие "Пропустить" включено в верхней правой части области задач.](../images/add-in-FRE-step-1.png)

*Рис. 2. Свести к минимуму количество экранов карусель только до того, что необходимо для эффективного сообщения вашего сообщения*

![Иллюстрация, показывающая шаг 2 карусели в первом запуске области задач Office настольного приложения. В этом примере в области задач имеется 3 экрана карусели.](../images/add-in-FRE-step-2.png)

*Рис. 3. Предоставление четкого вызова действий для выхода из первого запуска*

![Иллюстрация, показывающая шаг 3 карусели в первом запуске области задач Office настольного приложения. В этом примере на третьем и заключительном экране области задач показана кнопка для начала работы.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a>Представление ценности

Представление ценности — это ценностное предложение вашей надстройки: размещение логотипа, ясно сформулированное ценностное предложение, краткое описание или обзор функций, а также призыв к действию.

*Рис. 4. Placemat значения с логотипом, предложением четкого значения, сводка функций и вызов к действию*

![Иллюстрация, показывающая placemat значения в первом опытом запуска Office области задач настольных приложений. В этом примере на области задач отображается логотип надстройки, описание надстройки и кнопка для запуска.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a>Представление видео

Представление видео показывает пользователям видеоролик перед тем, как они начнут использовать вашу надстройку.

*Рис. 5. Первый запуск видео-placemat — экран содержит изображение из видео с кнопкой воспроизведения и кнопкой "Вызов к действию"*

![Иллюстрация, показывающая видео-placemat в первом опытом запуска Office области задач настольного приложения.](../images/add-in-FRE-video.png)

*Рис. 6. Video player — Пользователи, представленные с видео в диалоговом окне*

![Иллюстрация, показывающая видео в диалоговом окне с Office настольного приложения и области задач надстройки в фоновом режиме.](../images/add-in-FRE-video-dialog.png)
