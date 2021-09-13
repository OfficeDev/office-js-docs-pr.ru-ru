---
title: Настройка среды разработки
description: Настройка среды разработчика для создания Office надстройки.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 71982a51e4941cb90a488f317cf6f771ccf5b005
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150853"
---
# <a name="set-up-your-development-environment"></a>Настройка среды разработки

В этом руководстве вы можете настроить инструменты, чтобы Office надстройки, следуя нашим быстрым стартам или учебникам. Необходимо установить инструменты из приведенного ниже списка. Если у вас уже установлены эти установки, вы готовы начать быстрое начало, например, [Excel React.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Учетная Microsoft 365, которая включает версию подписки Office
- Редактор кода по вашему выбору

В этом руководстве предполагается, что вы знаете, как использовать средство командной строки.

## <a name="install-nodejs"></a>Установите Node.js.

Node.js является временем запуска JavaScript, необходимое для разработки Office надстройки.

Установите [Node.js, скачав последнюю рекомендуемую версию с веб-сайта.](https://nodejs.org) Следуйте инструкциям по установке операционной системы.

## <a name="install-npm"></a>Установка npm

npm — это реестр программного обеспечения с открытым исходным кодом, из которого можно скачать пакеты, используемые Office надстройки.

Чтобы установить npm, запустите следующую строку в командной строке.

```command&nbsp;line
    npm install npm -g
```

Чтобы проверить, установлена ли у вас npm и установлена версия, запустите следующую строку в командной строке.

```command&nbsp;line
npm -v
```

Может потребоваться использовать диспетчер версий node, чтобы разрешить переключаться между несколькими версиями Node.js npm, но это не является строго необходимым. Сведения о том, как это сделать, см. в инструкциях [npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## <a name="get-microsoft-365"></a>Получить Microsoft 365

Если у вас еще нет Microsoft 365 учетной записи, вы можете получить бесплатную 90-дневную возобновляемую подписку Microsoft 365, которая включает все Office приложения, присоединившись к программе [Microsoft 365](https://developer.microsoft.com/office/dev-program)разработчика .

## <a name="install-a-code-editor"></a>Установка редактора кода

Для создания веб-частей можно использовать любой редактор кода или интерфейс IDE, поддерживающий клиентскую разработку, например:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io);
- [Webstorm](https://www.jetbrains.com/webstorm).

## <a name="next-steps"></a>Следующие шаги

Попробуйте создать собственную надстройку или использовать Script Lab, чтобы попробовать встроенные образцы.

### <a name="create-an-office-add-in"></a>Создание надстройки Office

Вы можете быстро создать простую надстройку для Excel, OneNote, Outlook, PowerPoint, Project или Word с помощью [5-минутного краткого руководства по началу работы](../index.yml). Если вы уже ознакомились с кратким руководством и хотите создать более сложную надстройку, воспользуйтесь [учебником](../index.yml).

### <a name="explore-the-apis-with-script-lab"></a>Изучение API с помощью Script Lab

Изучите библиотеку встроенных примеров в [Script Lab](explore-with-script-lab.md), чтобы ознакомиться с возможностями API JavaScript для Office.

## <a name="see-also"></a>См. также

- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Проектирование надстроек Office](../design/add-in-design.md)
- [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
- [Публикация надстроек Office](../publish/publish.md)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)