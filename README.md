# Экспорт таблиц в виде отчетностей PDF и XLS

####Использование по умолчанию:
```html
<nf-document-export-button label="Экспорт" control="{{control}}"></nf-document-export-button>
```

####Использование с шаблоном (шапкой) при выгрузке:
```html
<nf-document-export-button label="Экспорт" control="{{control}}">
    <nf-document-export-template slot="header" type="title" align="center" size="h2" label="Отчет за период"></nf-document-export-template>
    <nf-document-export-template slot="header" type="new-line"></nf-document-export-template>
    <nf-document-export-template slot="header" type="filter" label="Наименование субъекта РФ" text="Москва"></nf-document-export-template>
    <nf-document-export-template slot="footer" type="filter" label="Период" text="01.01.2021-01.03.2021"></nf-document-export-template>
</nf-document-export-button>                        
```
###Аттрибуты:

####type
```text
Определяет тип шаблона
```

- title - формат заголовок (жирный текст+по центру)
- new-line - формирования новой линии
- filter - для фильтров, при отсутствующем параметре text не отображается

####align
```text
позиционирование текста
```

- left
- right
- center 

size
```text
Размер шрифта, только для type="title"
```
- H1
- H2
- H3
- H4

####label
```text
Наименование, обязательный
```

####text
```text
Текст
```

#### slot
```text
Размещение текста
```

- header
- footer
