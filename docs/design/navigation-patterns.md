---
title: Шаблоны навигации для надстроек Office
description: Узнайте, как использовать командные полосы, вкладки и кнопки назад для разработки навигации Office надстройки.
ms.date: 06/26/2018
ms.localizationpriority: medium
ms.openlocfilehash: dc7d75c9e914cf6294409590783e5ef73670dcc5
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743225"
---
# <a name="navigation-patterns"></a>Шаблоны навигации

Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную область экрана. Важно, чтобы навигация была интуитивно понятной, обеспечивала контекст и позволяла пользователю легко перемещаться по всей надстройке.

## <a name="best-practices"></a>Рекомендации

| Правильно    | Неправильно |
| :---- | :---- |
| Убедитесь, что пользователю доступен хорошо видимый параметр навигации. | Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.
| Используйте, по возможности, указанные ниже компоненты, позволяющие пользователям перемещаться по вашей надстройке. | Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке

## <a name="command-bar"></a>Панель команд

Командная панель — это поверхность в области задач, где находятся команды, которые работают на содержимом окна, панели или родительского региона, на которое оно расположено выше. Дополнительные функции включают точку доступа к меню "гамбургер", поиск и боковые команды.

![Иллюстрация, показывающая панели команд в области задач Office настольного приложения. В этом примере показана командная планка непосредственно под именем надстройки, которая включает меню гамбургера и поиск.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>Панель вкладок

На панели вкладок показана навигация с помощью кнопок с вертикально сложенным текстом и значками. Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.

![Иллюстрация, показывающая планку вкладок в Office области задач настольного приложения. В этом примере показана планка вкладок непосредственно под именем надстройки с вкладками "Главная", "Параметры", "Избранное" и "Учетная запись".](../images/add-in-tab-bar.png)

## <a name="back-button"></a>Кнопка "Назад"

Кнопка "Назад" позволяет пользователям восстановиться после навигационного действия, нажатого на сверлом. Этот шаблон помогает пользователям следовать упорядоченной последовательности действий.

![Иллюстрация, показывающая кнопку назад в области задач Office настольного приложения. В этом примере показана кнопка "Назад" сразу под именем надстройки в верхнем левом ок.](../images/add-in-back-button.png)
