---
title: Настройка среды разработки
description: Настройка среды разработчика для создания надстройки Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: eddf8bdf7b20a54667e6f8eb38bdace801ea1813
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839714"
---
# <a name="set-up-your-development-environment"></a>Настройка среды разработки

Это руководство поможет вам настроить средства для создания надстройки Office, следуя нашим кратким руководствам или руководствам. Вам потребуется установить средства из приведенного ниже списка. Если у вас уже установлены эти приложения, вы можете начать быстрое начало работы, например, это краткое начало [Excel React.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Учетная запись Microsoft 365, включаемая версию Office по подписке
- Редактор кода по вашему выбору

В этом руководстве предполагается, что вы знаете, как использовать средство командной строки. 

## <a name="install-nodejs"></a>Установите Node.js.

Node.js является среде запуска JavaScript, необходимо разрабатывать современные надстройки Office.

Установите Node.js, [скачав последнюю рекомендуемую версию с веб-сайта.](https://nodejs.org) Следуйте инструкциям по установке операционной системы.

## <a name="install-npm"></a>Установка npm

npm — это реестр программного обеспечения с открытым кодом, из которого можно скачать пакеты, используемые при разработке надстройки Office.

Чтобы установить npm, в командной строке запустите следующую команду:

```command&nbsp;line
    npm install npm -g
```

Чтобы проверить, установлен ли npm, и увидеть установленную версию, в командной строке запустите следующую команду:

```command&nbsp;line
npm -v
```

Вы можете использовать диспетчер версий Node, чтобы разрешить переключение между несколькими версиями Node.js npm, но это не является строго необходимым. Подробные сведения о том, как это сделать, см. в [инструкциях npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## <a name="get-office-365"></a>Получить Office 365

Если у вас еще нет учетной записи Microsoft 365, вы можете оформить бесплатную возобновляемую подписку на Microsoft 365 на 90 дней, присоединившись к [программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-a-code-editor"></a>Установка редактора кода

Для создания веб-частей можно использовать любой редактор кода или интерфейс IDE, поддерживающий клиентскую разработку, например:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io);
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Дальнейшие действия

Попробуйте создать собственную надстройку или воспользуйтесь Script Lab, чтобы попробовать встроенные примеры.

### <a name="create-an-office-add-in"></a>Создание надстройки Office

Вы можете быстро создать простую надстройку для Excel, OneNote, Outlook, PowerPoint, Project или Word с помощью [5-минутного краткого руководства по началу работы](../index.yml). Если вы уже ознакомились с кратким руководством и хотите создать более сложную надстройку, воспользуйтесь [учебником](../index.yml).

### <a name="explore-the-apis-with-script-lab"></a>Изучение API с помощью Script Lab

Изучите библиотеку встроенных примеров в [Script Lab](explore-with-script-lab.md), чтобы ознакомиться с возможностями API JavaScript для Office.

## <a name="see-also"></a>См. также

- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
- [Разработка надстройки Office](../develop/develop-overview.md)
- [Проектирование надстроек Office](../design/add-in-design.md)
- [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
- [Публикация надстроек Office](../publish/publish.md)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)