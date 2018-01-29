
# <a name="guidelines-for-creating-labs-for-mix-using-labsjs"></a>Рекомендации по созданию лаборатории для Office Mix с помощью LabsJS



Библиотека LabsJS (labs.js) позволяет создавать специальные Надстройки Office (называемые лабораториями), которые интегрируются с Office Mix. После создания лаборатории обрабатываются в Office Mix с помощью Microsoft PowerPoint. Хотя эти компоненты называются "лабораториями", фактически мы создаем специальные Надстройки Office, а именно Надстройки Office Mix.

Содержимое LabsJS позволяет применять API JavaScript labs.js, обеспечивая инструкции и примеры. Эта библиотека создана на основе [API JavaScript для Office](http://dev.office.com/reference/add-ins/javascript-api-for-office) (Office.js) и обеспечивает уровень абстракции, оптимизированный для надстроек, внедренных в Office Mix.


## <a name="general-guidelines"></a>Общие рекомендации


Ниже приведены общие рекомендации, которые могут помочь при написании надстроек с использованием API LabJS.


### <a name="scripts"></a>Скрипты

Поскольку библиотека labs.js — это уровень абстракции библиотеки office.js и, следовательно, зависит от нее, в проект разработки необходимо включить оба файла библиотеки — office.js и labs.js. 

Библиотека office.js доступна по следующей ссылке: `<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>`.

Библиотека labs.js включена в пакет SDK LabsJS. Кроме того, она доступна в сети доставки содержимого (CDN) по адресу <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Обратите внимание, что рабочая версия вашей лаборатории должна ссылаться на версию в сети CDN.


 >**Примечание.** Помимо файла JavaScript (labs-1.0.4.js), мы предоставляем файл определения TypeScript, содержащий API лабораторий (labs-1.0.4.d.ts). Файл определения создан на основе TypeScript версии 0.9.1.1.


### <a name="callbacks-and-error-handling"></a>Обратные вызовы и обработка ошибок

Ряд методов в API labs.js работает асинхронно. Для выполнения асинхронных операций API принимает стандартный интерфейс обратного вызова  **ILabCallback**. 


```js
function(err, result) {
}
```

Метод обратного вызова принимает два параметра:  _err_ и _result_. Если ошибки нет, поле  _err_ сохраняет значение **null**. Поле  _result_ возвращает результат операции.

Операция обратного вызова никогда не запускается сразу, даже если получен немедленный результат. Вместо этого она запускается при отдельном выполнении цикла обработки событий JavaScript (посредством вызова  **setTimeout**). Используя это определение обратного вызова, вы можете легко интегрировать labs.js с выбранным API обещаний. Например, можно заменить эти обратные вызовы обещаниями jQuery простым методом преобразования, как показано в примере ниже.




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### <a name="lab-host-and-defaultlabhost"></a>Узел лаборатории и DefaultLabHost

Узел лаборатории (**ILabHost**) служит основой для разработки лабораторий. По умолчанию задан узел, интегрируемый с office.js.

Для тестирования и запуска лаборатории в labhost.html необходимо переключиться на узел, работающий в имитационной среде. В следующем примере кода показано, как это сделать с помощью параметра запроса. Кроме того, можно изменить  **DefaultHostBuilder**, чтобы интегрировать лабораторию с другой средой.




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### <a name="initialization"></a>Инициализация

С помощью инициализации устанавливается путь обмена данными между лабораторией и узлом. Инициализируйте лабораторию, вызвав следующий метод:


```js
Labs.connect((err, connectionResponse) => {});
```

После инициализации можно вызвать другие методы API labs.js. Параметр  _connectionResponse_ содержит сведения об узле, пользователе, а также другую информацию о подключении. Дополнительные сведения о возвращаемых значениях см. в разделе [Labs.Core.IConnectionResponse](http://dev.office.com/reference/add-ins/office-mix/labs.core.iconnectionresponse).


### <a name="time-format"></a>Формат времени

В Labs.js хранятся числа, представляющие миллисекунды, истекшие с 1 января 1970 г. в формате UTC. Это соответствует формату JavaScript [объекта Date](http://msdn.microsoft.com/ru-ru/library/ie/cd9w2te4%28v=vs.94%29.aspx),


### <a name="timeline"></a>Временная шкала

Кроме того, лаборатория может взаимодействовать с временной шкалой проигрывателя уроков. С ее помощью лаборатория определяет, когда проигрывателю уроков необходимо переходить к следующему слайду. Объект временной шкалы можно получить, вызвав метод  **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="handling-events"></a>Обработка событий


API событий LabsJS отслеживает события, связанные с лабораторией, и позволяет добавлять обработчики событий, чтобы иметь возможность отвечать на события или действовать в соответствии с ними. Объект  **EventTypes** включает три метода событий: **ModeChanged**,  **Activate** и **Deactivate**. 


### <a name="mode-change"></a>Смена режимов

Событие  **ModeChanged** запускается, когда определенная лаборатория переключается с режима редактирования на режим просмотра. Режим редактирования отображается, когда лаборатория просматривается в режиме редактирования PowerPoint. Режим просмотра отображается, когда в PowerPoint выполняется слайд-шоу или когда лаборатория отображается в проигрывателе уроков Office Mix. В режиме просмотра всегда должно отображаться то, что пользователь видит при выполнении лаборатории. В режиме редактирования пользователь может настраивать лабораторию.

Данные в объекте  **ModeChangedEventData**, передаваемые в обратный вызов, содержат информацию о текущем режиме. В следующем коде показано, как использовать событие  **ModeChanged**.




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### <a name="activate"></a>Activate

Событие  **activate** запускается, когда слайд PowerPoint, на котором находится лаборатория, становится активным в проигрывателе уроков.


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### <a name="deactivate"></a>Deactivate

Событие  **deactivate** запускается, когда слайд PowerPoint, на котором находится лаборатория, больше не активен.


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### <a name="timeline"></a>Временная шкала

Лаборатория может взаимодействовать с временной шкалой проигрывателя уроков. С ее помощью лаборатория определяет, когда проигрывателю уроков необходимо переходить к следующему слайду. Объект временной шкалы можно получить, вызвав метод  **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="additional-resources"></a>Дополнительные ресурсы



- [Надстройки Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
