---
title: Среды выполнения в надстройки Office
description: Сведения о средах выполнения, используемых надстройки Office.
ms.date: 09/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: c20845e6df6adfa2fa382f10e8c7f5de786aeab8
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467232"
---
# <a name="runtimes-in-office-add-ins"></a>Среды выполнения в надстройки Office

Надстройки Office выполняются в средах выполнения, внедренных в Office. В качестве интерпретируемого языка JavaScript должен выполняться в среде выполнения JavaScript. [Node.js](https://nodejs.org) и современные браузеры являются примерами таких сред выполнения. 

## <a name="types-of-runtimes"></a>Типы сред выполнения

Надстройки Office могут использовать два типа сред выполнения:

- Среда выполнения только для **JavaScript**: подсистема JavaScript, дополненная поддержкой [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API), полной [CORS (](https://developer.mozilla.org/docs/Web/HTTP/CORS)совместное использование ресурсов независимо от источника) и клиентского хранилища данных. Он не поддерживает [локальное хранилище или](https://developer.mozilla.org/docs/Web/API/Window/localStorage) файлы cookie.
- **Среда выполнения браузера**: включает все функции среды выполнения только для JavaScript и добавляет поддержку локального [хранилища,](https://developer.mozilla.org/docs/Web/API/Window/localStorage) обработчика отрисовки, который отображает HTML и файлы cookie.[](https://developer.mozilla.org/docs/Glossary/Rendering_engine)

Дополнительные сведения об этих типах см. далее в этой статье в среде выполнения [только для JavaScript](#javascript-only-runtime) и [в среде выполнения Браузера](#browser-runtime).

В следующей таблице показано, какие возможные функции надстройки используют каждый тип среды выполнения. 

| Тип среды выполнения | Функция надстройки |
|:-----|:-----|
| Только JavaScript | [Пользовательские функции](../excel/custom-functions-overview.md) Excel</br>(за исключением случаев, когда [среда выполнения является](#shared-runtime) общей или надстройка работает в Office в Интернете)</br></br>[Задача на основе событий Outlook](../outlook/autolaunch.md)</br>(только если надстройка работает в Outlook для Windows)|
| Обозреватель | [область задач](../design/task-pane-add-ins.md)</br></br>[диалоговое окно](../develop/dialog-api-in-office-add-ins.md)</br></br>[команда function](../design/add-in-commands.md#types-of-add-in-commands)</br></br>[Пользовательские функции](../excel/custom-functions-overview.md) Excel</br>(если среда [выполнения является общей](#shared-runtime) или надстройка работает в Office в Интернете)</br></br>[Задача на основе событий Outlook](../outlook/autolaunch.md)</br>(если надстройка работает в Outlook для Mac или Outlook в Интернете)|

В следующей таблице показаны те же сведения, упорядоченные по типу среды выполнения, которые используются для различных возможных функций надстройки.

| Функция надстройки | Тип среды выполнения в Windows | Тип среды выполнения на Компьютере Mac | Тип среды выполнения в Интернете |
|:-----|:-----|:-----|:-----|
|Пользовательские функции Excel | Только JavaScript</br>(но *браузер* при совместном использовании среды выполнения)|Только JavaScript</br>(но *браузер* при совместном использовании среды выполнения)| Обозреватель |
|Задачи на основе событий Outlook | Только JavaScript | Обозреватель | Обозреватель |
|Надстройки области задач | Обозреватель | Обозреватель | Обозреватель |
|диалоговое окно | Обозреватель | Обозреватель | Обозреватель |
|команда function | Обозреватель | Обозреватель | Обозреватель |


В Office в Интернете все всегда выполняется в среде выполнения типа браузера. На самом деле, за одним исключением, все содержимое надстройки в Интернете выполняется  в одном браузере: в браузере, в котором пользователь Office в Интернете. Исключением является то, что диалоговое окно открывается с помощью вызова [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) и параметр [DialogOptions.displayInIFrame](/javascript/api/office/office.dialogoptions#office-office-dialogoptions-displayiniframe-member)  `true`не передается и не задается в значение . Если параметр не передается ( `false` поэтому он имеет значение по умолчанию), диалоговое окно открывается в собственном процессе. Тот же принцип применяется к методу [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) и [параметру OfficeRuntime.DisplayWebDialogOptions.displayInIFrame](/javascript/api/office-runtime/officeruntime.displaywebdialogoptions#office-runtime-officeruntime-displaywebdialogoptions-displayiniframe-member) .

При запуске надстройки на платформе, отличной от веб-платформы, применяются следующие принципы.

- Диалоговое окно запускается в собственном процессе среды выполнения. 
- Задача на основе событий Outlook выполняется в собственном процессе среды выполнения. 
- По умолчанию области задач, команды функций и пользовательские функции Excel выполняются в собственном процессе среды выполнения. Однако для некоторых основных приложений Office манифест надстройки можно настроить таким образом, чтобы все два или все три приложения могли выполняться в одной среде выполнения. См [. раздел "Общая среда выполнения"](#shared-runtime).

В зависимости от ведущего приложения Office и функций, используемых в надстройке, в надстройке может быть много сред выполнения. Каждый из них обычно выполняется в собственном процессе, но не обязательно одновременно. Ниже приведены примеры.

- Надстройка PowerPoint или Word, которая не предоставляет общий доступ к средам выполнения и включает следующие функции, имеет до трех сред выполнения.

  - Область задач
  - Команда функции
  - Диалоговое окно (диалоговое окно можно запустить из области задач или из команды функции.) 
  
      > [!NOTE]
      > Не рекомендуется одновременно открывать несколько диалогов, но если надстройка позволяет пользователю открывать один из них из области задач, а другой из команды функции одновременно, эта надстройка будет иметь четыре среды выполнения. Область задач и заданный вызов команды функции могут иметь только одно открытое диалоговое окно за раз; Но если команда функции вызывается несколько раз, новый диалог открывается поверх предшественницы с каждым вызовом, поэтому может быть много сред выполнения. В оставшейся части этого списка игнорируется возможность нескольких открытых диалогов.

- Надстройка Excel, которая не предоставляет общий доступ к средам выполнения и включает следующие функции, имеет до *четырех* сред выполнения.

  - Область задач
  - Команда функции
  - Пользовательская функция
  - Диалоговое окно (диалоговое окно можно запустить из области задач, команды функции или пользовательской функции.)

- Надстройка Excel с одинаковыми функциями и настроена для совместного использования одной и той же среды выполнения в области задач, команде функции и настраиваемой  функции, имеет две среды выполнения. Общая среда выполнения может открывать только один диалог за раз.
- Надстройка Excel с одинаковыми функциями, за исключением того, что она не имеет диалогового окна и настроена для совместного использования одной и той же среды выполнения в области задач, команде функции  и настраиваемой функции, имеет одну среду выполнения.
- Надстройка Outlook со следующими функциями имеет до *четырех* сред выполнения. (Среды выполнения не могут совместно использоваться в Outlook.)

  - Область задач
  - Команда функции
  - Задача на основе событий
  - Диалоговое окно (диалоговое окно можно запустить из области задач или команды функции, но не из задачи на основе событий.)

## <a name="share-data-across-runtimes"></a>Совместное использование данных в средах выполнения

> [!NOTE]
> - Если вы знаете `displayInIFrame` `true`, что надстройка будет использоваться только в Office в Интернете и что она не будет открывать диалоги с заданной для параметра настройкой, этот раздел можно пропустить. Так как все в вашей надстройке выполняется в одном процессе среды выполнения, можно просто использовать глобальные переменные для совместного использования данных между компонентами.
> - Как отмечалось выше [в типах сред выполнения](#types-of-runtimes), тип среды выполнения, используемой компонентом, частично зависит от платформы. Рекомендуется не использовать код надстройки, ветвления которого основаны на платформе, поэтому в этом разделе рекомендуется использовать методы, которые будут работать на разных платформах. Ниже указан только один случай, в котором требуется код ветвления. 

Для надстроек Excel, PowerPoint и Word используйте общую среду выполнения[](#shared-runtime), если для общего доступа к данным требуются две или более функций, за исключением диалоговых окон. В Outlook или сценариях, где совместное использование среды выполнения невозможно, требуются альтернативные методы. Части надстройки, которые находятся в отдельных процессах среды выполнения, не совместно используют глобальные данные автоматически и обрабатываются сервером веб-приложений надстройки как отдельные сеансы, поэтому [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) нельзя использовать для совместного использования данных между ними. *В следующем руководстве предполагается, что вы не используете общую среду выполнения.*

- Передача данных между диалогом и родительской областью задач, командой функции или пользовательской функцией с помощью методов [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) и [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) . 

    > [!NOTE]
    > Методы `OfficeRuntime.storage` не могут вызываться в диалоговом окне, поэтому это не вариант для совместного использования данных между диалогом и другой средой выполнения. 

- Чтобы совместно использовать данные между областью задач и командой функции, сохраните данные в [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), который используется во всех средах выполнения, которые имеют доступ к одному и тем же [источнику](https://developer.mozilla.org/docs/Glossary/Origin). 
    > [!NOTE]
    > LocalStorage недоступен в среде выполнения только для JavaScript, поэтому он недоступен в пользовательских функциях Excel. Его также нельзя использовать для совместного использования данных с задачами на основе событий Outlook (так как эти задачи используют среду выполнения только javaScript на некоторых платформах).

    > [!TIP]
    > Данные, `Window.localStorage` хранимые между сеансами надстройки, совместно используются надстройки с одинаковым источником. Обе эти характеристики часто нежелательны для надстройки. 
    >
    > - Чтобы каждый сеанс данной надстройки запускал новый вызов метода [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) при запуске надстройки. 
    > - Чтобы разрешить сохранение некоторых хранимых значений, но повторную инициализацию других значений, используйте [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) при запуске надстройки для каждого элемента, который должен быть сброшен до начального значения. 
    > - Чтобы полностью удалить элемент, вызовите [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem).

- Чтобы совместно использовать данные между пользовательской функцией Excel и любой другой средой выполнения, используйте [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage).
- Чтобы совместно использовать данные между задачей на основе событий Outlook и командой области задач или функцией, необходимо ветвление кода по значению свойства [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) . 

    - Если значение равно `PC` (Windows), храните и извлекайте данные с помощью API [Office.sessionData](/javascript/api/outlook/office.sessiondata) .
    - Если это значение, `Mac`используйте `Window.localStorage` его, как описано выше в этом списке.

К другим способам совместного использования данных относятся следующие:

- Храните общие данные в оперативной базе данных, доступной для всех сред выполнения.
- Храните общие данные в файле cookie для домена надстройки, чтобы поделиться этими данными между средами выполнения браузера. Среды выполнения только для JavaScript не поддерживают файлы cookie.

Дополнительные сведения см. в разделе ["Сохранение](../develop/persisting-add-in-state-and-settings.md) состояния и параметров надстройки" и "Управление состоянием и параметрами надстройки [Outlook"](../outlook/manage-state-and-settings-outlook.md).

## <a name="javascript-only-runtime"></a>Среда выполнения только для JavaScript

Среда выполнения только для JavaScript, используемая в надстройки Office, представляет собой изменение среды открытый код, изначально созданной для [React Native.](https://reactnative.dev/) Он содержит подсистему JavaScript, дополненную поддержкой [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [Full CORS (](https://developer.mozilla.org/docs/Web/HTTP/CORS)общий доступ к ресурсам независимо от источника) и [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage). У него нет обработчика отрисовки и он не поддерживает файлы cookie или [локальное хранилище](https://developer.mozilla.org/docs/Web/API/Window/localStorage). 

Этот тип среды выполнения используется в задачах На основе событий Outlook только в Office для Windows и в пользовательских функциях  Excel, за исключением случаев, когда пользовательские функции совместно [работают со средой выполнения](#shared-runtime). 

- При использовании для пользовательской функции Excel среда выполнения запускается при пересчете листа или вычислении пользовательской функции. Она не завершает работу, пока книга не будет закрыта.  
- При использовании в задаче на основе событий Outlook среда выполнения запускается при возникновении события. Он заканчивается, когда происходит первое из следующих действий.

  - Обработчик событий вызывает метод `completed` своего параметра события.
  - С момента запуска события прошло 5 минут.
  - Пользователь изменяет фокус с окна, в котором было активируется событие, например окно создания сообщения.

Среда выполнения JavaScript использует меньше памяти и запускается быстрее, чем среда выполнения браузера, но имеет меньше функций.

## <a name="browser-runtime"></a>Среда выполнения браузера

Надстройки Office используют другую среду выполнения браузера в зависимости от платформы, на которой работает Office (Веб, Mac или Windows), а также от версии и сборки Windows и Office. Например, если пользователь выполняет Office в Интернете в браузере FireFox, используется среда выполнения Firefox. Если пользователь работает под управлением Office на Mac, используется среда выполнения Safari. Если пользователь работает с Office в Windows, то среда выполнения предоставляется в Edge или Internet Explorer в зависимости от версии Windows и Office. Подробные сведения можно найти в [браузерах, используемых надстройки Office](../concepts/browsers-used-by-office-web-add-ins.md).

Все эти среды выполнения включают обработчик отрисовки HTML и обеспечивают поддержку [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API), полной [CORS (](https://developer.mozilla.org/docs/Web/HTTP/CORS)общего доступа к ресурсам независимо от источника[](https://developer.mozilla.org/docs/Web/API/Window/localStorage)) и локального хранилища и файлов cookie.

Жизненный цикл среды выполнения браузера зависит от реализуемой функции и от того, предоставляется ли к ней общий доступ.

- При запуске надстройки с областью задач запускается среда выполнения браузера, если только это не общая среда выполнения, которая уже запущена. Если это общая среда выполнения, она завершает работу при закрытии документа. Если это не общая среда выполнения, она завершает работу при закрытии области задач.
- При открытии диалогового окна запускается среда выполнения браузера. Он завершает работу при закрытии диалогового окна.
- При выполнении команды функции (которая происходит, когда пользователь нажмет кнопку или пункт меню), запускается среда выполнения браузера, если только она не является общей средой выполнения, которая уже запущена. Если это общая среда выполнения, она завершает работу при закрытии документа. Если это не общая среда выполнения, она завершает работу, когда происходит первое из следующих действий.
 
  - Команда функции вызывает метод `completed` своего параметра события.
  - С момента запуска события прошло 5 минут. (Если диалоговое окно было открыто в команде функции и по-прежнему открыто, когда истекло время ожидания родительской среды выполнения, среда выполнения диалогового окна остается запущенной до тех пор, пока диалоговое окно не будет закрыто.)

- Когда пользовательская функция Excel использует общую среду выполнения, среда выполнения типа браузера запускается, когда пользовательская функция вычисляет, если общая среда выполнения еще не запущена по какой-либо другой причине. Он завершает работу при закрытии документа.

> [!NOTE]
> При совместном использовании среды выполнения [](#shared-runtime)код может закрыть область задач без завершения работы надстройки. [Дополнительные сведения см](../develop/show-hide-add-in.md). в разделе "Показать или скрыть область задач" надстройки Office.

Среда выполнения браузера имеет больше функций, чем среда выполнения только для JavaScript, но запускается медленнее и использует больше памяти.

### <a name="shared-runtime"></a>Общее время выполнения

Общая среда выполнения не является типом среды выполнения. Она относится к [среде](#browser-runtime) выполнения типа браузера, которая предоставляется функциями надстройки, которые в противном случае будут иметь собственную среду выполнения. В частности, вы можете настроить область задач надстройки и команды функций для совместного использования среды выполнения. В надстройке Excel можно также настроить пользовательские функции для совместного использования среды выполнения области задач или команды функции или и того, и другое. При этом пользовательские функции выполняются в среде выполнения типа браузера, а не в среде выполнения только [для JavaScript](#javascript-only-runtime) , как в противном случае. Сведения [о](../develop/configure-your-add-in-to-use-a-shared-runtime.md) преимуществах и ограничениях общего доступа к средам выполнения и инструкции по настройке надстройки для использования общей среды выполнения см. в разделе "Настройка надстройки для использования общей среды выполнения". Вкратце, среда выполнения только для JavaScript использует меньше памяти и запускается быстрее, но имеет меньше функций.

> [!NOTE]
> - Делиться средами выполнения можно только в Excel, PowerPoint и Word. 
> - Невозможно настроить диалоговое окно для совместного использования среды выполнения. Каждое диалоговое окно всегда имеет свой собственный, за исключением случаев, когда диалог запускается в Office в Интернете `displayInIFrame` с параметром, для которых задано значение `true`.
> - Общая среда выполнения никогда не использует исходную среду выполнения Microsoft Edge WebView (EdgeHTML). Если выполняются условия использования Microsoft Edge с WebView2 (Chromium на основе) (как указано в браузерах, используемых надстройки [Office](../concepts/browsers-used-by-office-web-add-ins.md)), используется эта среда выполнения. В противном случае используется среда выполнения Internet Explorer 11.