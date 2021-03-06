---
title: Шаблоны навигации для надстроек Office
description: Ознакомьтесь с рекомендациями по использованию панелей команд, вкладок и кнопок "назад", чтобы разработать навигацию для надстройки Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132034"
---
# <a name="navigation-patterns"></a>Шаблоны навигации

Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную область экрана. Важно, чтобы навигация была интуитивно понятной, обеспечивала контекст и позволяла пользователю легко перемещаться по всей надстройке.

## <a name="best-practices"></a>Рекомендации

| Правильно    | Неправильно |
| :---- | :---- |
| Убедитесь, что пользователю доступен хорошо видимый параметр навигации. | Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.
| Используйте, по возможности, указанные ниже компоненты, позволяющие пользователям перемещаться по вашей надстройке. | Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке

## <a name="command-bar"></a>Панель команд

Панель элементов управления — это поверхность области задач, в которой размещаются команды, работающие с содержимым окна, панели или родительской области, расположенной выше. Дополнительные функции включают точку доступа к меню "гамбургер", поиск и боковые команды.

![Иллюстрация, демонстрирующая панель команд в области задач приложения Office для настольных ПК. В этом примере показана Командная строка, расположенная сразу под именем надстройки, которая включает меню "гамбургер" и поиск.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>Панель вкладок

Панель вкладок показывает навигацию с помощью кнопок с вертикальным текстом и значками. Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.

![Иллюстрация, на которой показана панель вкладок в области задач приложения Office для настольных ПК. В этом примере отображается панель вкладок сразу под именем надстройки с вкладками "Домашняя страница", "Параметры", "Избранное" и "учетная запись".](../images/add-in-tab-bar.png)

## <a name="back-button"></a>Кнопка "Назад"

Кнопка "назад" позволяет пользователям восстанавливаться при переходе по навигации. Этот шаблон помогает пользователям следовать упорядоченной последовательности действий.

![Иллюстрация, на которой показана кнопка "назад" в области задач приложения Office для настольных ПК. В этом примере показана кнопка "назад", расположенная сразу под именем надстройки, в верхнем левом углу.](../images/add-in-back-button.png)
