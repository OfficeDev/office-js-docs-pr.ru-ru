---
title: Настройка среды разработки
description: Настройка среды разработчика для создания надстроек Office.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e03ea7f55786107354f9d5a92e0cb30ffb559ec
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616008"
---
# <a name="set-up-your-development-environment"></a>Настройка среды разработки

Это руководство поможет вам настроить средства для создания надстроек Office, следуя нашим кратким руководствам или руководствам. Если у вас уже установлены эти компоненты, вы можете приступить к работе с кратким React [excel.](../quickstarts/excel-quickstart-react.md)

## <a name="get-microsoft-365"></a>Получение Microsoft 365

Вам нужна учетная запись Microsoft 365. Вы можете получить бесплатную 90-дневную возобновляемую подписку на Microsoft 365, которая включает все приложения Office, присоединившись к программе для разработчиков [Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-the-environment"></a>Установка среды

Существует два типа сред разработки, которые можно выбрать. Шаблоны проектов надстроек Office, созданных в двух средах, отличаются, поэтому если над проектом надстройки будут работать несколько человек, все они должны использовать ту же среду. 

- **Node.js среды**: рекомендуется. В этой среде ваши средства устанавливаются и выполняются в командной строке. Серверная часть веб-приложения надстройки написана на JavaScript или TypeScript и размещается в среде Node.js выполнения. В этой среде существует множество полезных средств разработки надстроек, таких как подсистема анализатора Office и средство запуска пакетов и задач с именем WebPack. Средство создания и формирования шаблонов проекта Yo Office часто обновляется.
- **Среда Visual Studio**: выберите эту среду, только если компьютером разработки является Windows и вы хотите разработать серверную часть надстройки с использованием языка и платформы на основе .NET, например ASP.NET. Шаблоны проектов надстроек в Visual Studio обновляются не так часто, как в Node.js среде. Клиентский код нельзя отлаживать с помощью встроенного отладчика Visual Studio, но вы можете выполнять отладку клиентского кода с помощью средств разработки браузера. Дополнительные сведения см. на **вкладке среды Visual Studio** .

> [!NOTE]
> Visual Studio для Mac не включает шаблоны шаблонов проектов для надстроек Office, поэтому, если компьютером разработки является Компьютер Mac, следует работать с Node.js средой.

Выберите вкладку для выбранной среды. 

# <a name="nodejs-environment"></a>[Node.js среды](#tab/yeomangenerator)

Основные средства для установки:

- Node.js
- npm
- Редактор кода по вашему выбору
- Yo Office
- Анализатор JavaScript для Office

В этом руководстве предполагается, что вы знаете, как использовать средство командной строки.

### <a name="install-nodejs-and-npm"></a>Установка Node.js npm

Node.js — это среда выполнения JavaScript, используемая для разработки современных надстроек Office.

Установите Node.js [, скачав последнюю рекомендуемую версию с веб-сайта](https://nodejs.org). Следуйте инструкциям по установке операционной системы.

npm — это открытый код программного обеспечения, из которого скачиваются пакеты, используемые при разработке надстроек Office. Обычно он устанавливается автоматически при установке Node.js. Чтобы проверить, установлен ли npm, и просмотреть установленную версию, выполните следующую команду в командной строке.

```command&nbsp;line
npm -v
```

Если по какой-либо причине вы хотите установить его вручную, выполните следующую команду в командной строке.

```command&nbsp;line
npm install npm -g
```

> [!TIP]
> Вы можете использовать диспетчер версий Node, чтобы разрешить переключение между несколькими версиями Node.js npm, но это не является обязательным. Дополнительные сведения о том, как это сделать, см. в [инструкциях npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

### <a name="install-a-code-editor"></a>Установка редактора кода

Для создания веб-частей можно использовать любой редактор кода или интерфейс IDE, поддерживающий клиентскую разработку, например:

- [Visual Studio Code](https://code.visualstudio.com/) (рекомендуется)
- [Atom](https://atom.io);
- [Webstorm](https://www.jetbrains.com/webstorm)

### <a name="install-the-yeoman-generator-mdash-yo-office"></a>Установка генератора Yeoman &mdash; Yo Office

Средство создания и формирования шаблонов проектов — это генератор [Yeoman](../develop/yeoman-generator-overview.md) для надстроек Office, который известен как **Yo Office**. Необходимо установить последнюю версию [Yeoman](https://github.com/yeoman/yo) и Yo Office. Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### <a name="install-and-use-the-office-javascript-linter"></a>Установка и использование подстроки JavaScript для Office

Корпорация Майкрософт предоставляет анализатор Кода JavaScript, который помогает выявлять распространенные ошибки при использовании библиотеки JavaScript для Office. Чтобы установить модуль анализатора, выполните следующие две команды (после установки Node.js [npm](#install-nodejs-and-npm)).

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Если вы создаете проект надстройки Office с помощью генератора [Yeoman для надстроек Office](../develop/yeoman-generator-overview.md) , остальные настройки выполняются за вас. Запустите модуль анализатора с помощью следующей команды в терминале редактора, например Visual Studio Code или в командной строке. Проблемы, найденные анализатором кода, отображаются в терминале или запросе, а также непосредственно в коде при использовании редактора, который поддерживает сообщения анализатора, такие как Visual Studio Code. (Сведения об установке генератора Yeoman см. в разделе [генератора Yeoman для надстроек Office](../develop/yeoman-generator-overview.md).)

```command&nbsp;line
npm run lint
```

Если проект надстройки был создан другим способом, выполните следующие действия.

1. В корне проекта создайте текстовый файл с именем **.eslintrc.json**, если он еще не существует. Убедитесь, что у него есть свойства с именем `plugins` `extends`и оба типа массива. Массив `plugins` должен включаться, `"office-addins"` а `extends` массив должен включать .`"plugin:office-addins/recommended"` Ниже приведен простой пример. Файл **.eslintrc.json** может иметь дополнительные свойства и дополнительные элементы двух массивов.

   ```json
   {
     "plugins": [
       "office-addins"
     ],
     "extends": [
       "plugin:office-addins/recommended"
     ]
   }
   ```

1. В корне проекта откройте файл **package.json** `scripts` и убедитесь, что массив содержит следующий элемент.

   ```json
   "lint": "office-addin-lint check",
   ```

1. Запустите модуль анализатора с помощью следующей команды в терминале редактора, например Visual Studio Code или в командной строке. Проблемы, найденные анализатором кода, отображаются в терминале или запросе, а также непосредственно в коде при использовании редактора, который поддерживает сообщения анализатора, такие как Visual Studio Code.

   ```command&nbsp;line
   npm run lint
   ```

# <a name="visual-studio-environment"></a>[Среда Visual Studio](#tab/visualstudio)

### <a name="install-visual-studio"></a>Установка Visual Studio

Если у вас нет Visual Studio 2017 (для Windows) или более поздней версии, установите последнюю версию из [скачиваемых файлов Visual Studio](https://visualstudio.microsoft.com/downloads/). Обязательно включите рабочую **нагрузку разработки Office или SharePoint** , когда установщик предложит указать рабочие нагрузки. Другие рабочие нагрузки, которые могут потребоваться, — это средства веб-разработки для **.NET**, **поддержка языка JavaScript и TypeScript** (для написания кода на стороне клиента надстройки) и рабочие нагрузки, связанные ASP.NET.

> [!TIP]
> Начиная с июня 2022 г. XML-схемы манифеста надстройки Office, установленные с помощью Visual Studio, не являются последней версией. Это может повлиять на надстройки в зависимости от того, какие функции надстроек они используют. Поэтому может потребоваться обновить XML-схемы для манифеста. Дополнительные сведения см [. в статье об ошибках проверки схемы манифеста в проектах Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

> [!NOTE]
> Сведения об отладке клиентского кода при использовании среды Visual Studio см. в разделе "Отладка надстроек [Office" в Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md). Отладка кода на стороне сервера выполняется так же, как и любое веб-приложение, созданное в Visual Studio. См [. сведения на стороне клиента или на стороне сервера](../testing/debug-add-ins-overview.md#server-side-or-client-side).

---

## <a name="install-script-lab"></a>Установка Script Lab

Script Lab — это инструмент для быстрого создания прототипов кода, который вызывает API библиотеки JavaScript для Office. Script Lab является надстройка Office и может быть установлена из AppSource [Script Lab](https://appsource.microsoft.com/marketplace/apps?search=script%20lab&page=1). Существует версия для Excel, PowerPoint и Word, а также отдельная версия для Outlook. Сведения о том, как использовать Script Lab, см. в статье ["Обзор API JavaScript для Office с помощью Script Lab"](explore-with-script-lab.md).

## <a name="next-steps"></a>Дальнейшие действия

Попробуйте создать собственную надстройку [или Script Lab](explore-with-script-lab.md) использовать встроенные примеры.

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