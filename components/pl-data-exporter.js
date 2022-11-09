import dayjs from 'dayjs/esm/index.js';
import { PlElement, css } from "polylib";

class PlDataExporter extends PlElement {
    static properties = {
        exportTitle: {
            type: String
        },
        data: {
            type: Object
        },
        booleanNames: {
            type: Array,
            value: () => ['Да', 'Нет']
        },
        tryOpenInBrowser: {
            type: Boolean
        }
    };

    static css = css`
        :host {
            display: none;
        }
    `;

    /** Метод для получения данных
     * @private
     */
    async _getDataForExport(options) {
        let parentForm = this._findForm();
        let components = this._getBindedComponents(parentForm);
        let grid = components.find(component => component.tagName == 'PL-GRID');

        let filterItems = components
            .filter(component =>
                component.tagName == 'PL-FILTER-CONTAINER');

        let templateItems = Array.prototype.slice.call(this.querySelectorAll('PL-DOCUMENT-EXPORT-TEMPLATE'));

        let filters = this._getFilters(filterItems);
        let dataSet = components.find(component => component.tagName == 'PL-DATASET');
        let templates = this._getTemplates(templateItems);

        let fields = grid._columns.filter(c => !c.hidden).map(item => item.field);
        fields.unshift('counter');

        let widths = grid._columns.filter(c => !c.hidden).map(item => item.width || item._calculatedWidth);
        widths.unshift(40);

        let columns = grid._columns.filter(c => !c.hidden).map(item => item.header);
        columns.unshift('п/н');

        let data = await this._getFullData(grid, fields, dataSet, options);
        return {
            sheetName: parentForm.formTitle,
            data: data,
            names: columns,
            fields: fields,
            filters: filters,
            widths: widths,
            templates: templates
        }
    }

    /**
     * Метод для обработки элементов для формирования заголовков и фильтров документа
     * @param {Array<Object>} templateItems - Шаблоны
     * @return {Array<Object>}
     * @private
     */
    _getTemplates(templateItems) {
        let templates = [];

        templateItems.forEach((component) => {
            switch (component.getAttribute('type')) {
                case 'title': // Заголовки первого (label) + размер (size) h1,h2,h3
                    templates.push({
                        type: 'title',
                        label: component.getAttribute('label'),
                        size: component.getAttribute('size') ? component.getAttribute('size') : 'h1'
                    });
                    break;
                case 'new-line': // Новая строка
                    templates.push({
                        type: 'new-line'
                    });
                    break;
                case 'filter': // Для фильтров, у которых есть информация по фильтру
                    if (component.getAttribute('text')) {
                        templates.push({
                            type: 'filter',
                            label: component.getAttribute('label'),
                            text: component.getAttribute('text'),
                            align: component.getAttribute('align') ? component.getAttribute('align') : 'left',
                        });
                    }
                    break;
            }
        });

        return templates;
    }

    _getFilters(filterItems) {
        let filters = [];

        if (this.exportTitle) {
            filters.push({
                label: this.exportTitle
            });
        }

        // filterItems.forEach(item => {
        //     if(item.getSelectedFilterValues) {
        //         item.getSelectedFilterValues().forEach(filt => {
        //             if(filt.value instanceof Date) {
        //                 filt.value = dayjs(filt.value).format('DD.MM.YYYY')
        //             }
        //             filters.push(filt)
        //         });
        //     }
        // });

        filters.push({
            label: 'Дата выгрузки',
            value: dayjs(new Date()).format('DD.MM.YYYY')
        });

        return filters.map(item => `${item.label}` + (item.value ? `: ${item.value}` : ``)).join('\r\n');
    }

    /**
     * Метод для получения данных с сервера, сортировки и подготовки для отображения
     * @async
     * @private
     * @param {Object} grid - грид
     * @param {Array<String>} fields - массив наименований полей
     * @param {Object} dataSet - датасет для получения данных
     */
    async _getFullData(grid, fields, dataSet, options) {
        let counter = 1;
        const booleanNames = this.booleanNames;

        let data = null;
        if (grid.tree) {
            let preData = await dataSet.execute(dataSet.args, { partialData: false, returnOnly: true });
            data = (preData && this._sortData(preData.data, grid.key, grid.pkey)) || [];
        }
        else {
            let preData = await dataSet.execute(dataSet.args, { partialData: false, returnOnly: true });
            data = preData?.data || [];
        }

        let resultData = data.map(item => fields.reduce((res, key) => {
            let v = key ? item[key] : null;
            if (!v && key === 'counter') {
                res.push(counter++);
            }
            else if (v instanceof Date) {
                res.push(dayjs(v).format("DD.MM.YYYY"));
            }
            else if (typeof v === 'boolean') {
                res.push(v ? booleanNames[0] : booleanNames[1]);
            }
            else {
                if (v || v === 0) {
                    res.push(v);
                } else {
                    res.push('');
                }
            }
            return res;
        }, []));

        return resultData;
    }

    _getBindedComponents(form) {
        let dataPropName = form._ti.binds.find(x => x.type == 'prop' && x.node == this).depend[0];
        return form._ti.binds.filter(x => x.type == 'prop' && x.depend.includes(dataPropName) && x.node != this).map(x => x.node)
    }

    _sortData(data, key, parentKey) {
        let counter = 1;
        let parents = data.filter(item => item[parentKey] == null);
        let childs = data.filter(item => item[parentKey] != null);
        let result = [];
        parents.forEach(el => {
            el.counter = counter++;
            result.push(el);
            if (el._haschilds) {
                let directChilds = childs.filter(childEl => childEl[parentKey] == el[key]);
                let childCounter = 1;
                directChilds.forEach(dirChildEl => {
                    this._sortChildsRecursively(result, dirChildEl, childs, key, parentKey, el.counter, childCounter++);
                });
            }
        });

        return result;
    }


    _sortChildsRecursively(result, el, childs, key, parentKey, parentCounter, childCounter) {
        el.counter = parentCounter + '.' + childCounter;
        result.push(el);
        if (el._haschilds) {
            let directChilds = childs.filter(childEl => childEl[parentKey] == el[key]);

            let directChildCounter = 1;
            directChilds.forEach(dirChildEl => {
                this.sortElementsRecursive(result, dirChildEl, childs, key, parentKey, el.counter, directChildCounter++);
            });
        }
    }

    _findForm() {
        let form = this;
        while (form._ti == null) {
            form = this.parentNode.host;
        }
        return form;
    }

    /** 
     * Экспорт в XLS 
     */
    async exportToXls() {
        this.isWait = true;
        let prepData = await this._getDataForExport({ exportMode: 'xls' });
        let template = await fetch('/api/document_export/getTemplate');
        let templateWorkBook = await XlsxPopulate.fromDataAsync(await template.blob());
        let templateWorkSheet = templateWorkBook.sheet(0);

        prepData.widths.forEach((width, index) => {
            this._setColumnWidth(templateWorkSheet, index, width)
        });

        let filtersSectionStyleId = templateWorkSheet.find('[filterSection]')[0]._styleId;
        let headersSectionStyleId = templateWorkSheet.find('[headerSection]')[0]._styleId;
        let dataSectionStyleId = templateWorkSheet.find('[dataSection]')[0]._styleId;

        let rows = this._prepareFiltersSection(prepData.filters, prepData.names, templateWorkSheet, filtersSectionStyleId);

        rows = this._setTemplates(rows, prepData.templates, prepData.names, templateWorkSheet);

        rows = this._prepareHeadersSection(rows + 1, prepData.names, templateWorkSheet, headersSectionStyleId);

        this._prepareDataSection(rows, prepData.names, prepData.data, templateWorkSheet, dataSectionStyleId, prepData.templates.length == 0);

        let blob = await templateWorkBook.outputAsync();

        this.isWait = false;

        this.saveAs(blob, prepData.sheetName + ".xlsx");
    }

    saveAs(blob, name) {
        var url = window.URL.createObjectURL(blob);
        var a = document.createElement("a");
        document.body.appendChild(a);
        a.href = url;
        a.download = name;
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    }

    /**
     * Формирование шаблона (шапки) для выгрузки
     * @param {Number} rowCount - Номер строки с которой начинается работа
     * @param {Array<Object>} templates - Шаблоны
     * @param {Array<String>} names - массив наименований колонок
     * @param {*} workSheet - объект Worksheet
     * @returns {Number} number - Номер конечной строки
     * @private
     */
    _setTemplates(rowCount, templates, names, workSheet) {
        if (templates.length == 0) return rowCount;
        rowCount = 1;
        templates.forEach((item) => {
            let filterRange = workSheet.range(`A${rowCount}:${this._getColumnName(names.length - 1)}${rowCount}`);
            filterRange
                .style({})
                .style("border", {})
                .style("fill", 0);
            workSheet.row(rowCount).height(null);
            switch (item.type) {
                case 'title':
                    filterRange
                        .style("horizontalAlignment", "center")
                        .style("bold", true)
                        .style("fontSize", (item.size === 'H1' ? 12 :
                            (item.size === 'H2' ? 10 :
                                (item.size === 'H3' ? 9 :
                                    (item.size === 'H4' ? 8 : 12)))))
                        .value(item.label)
                        .merged(true);
                    rowCount++;
                    break;
                case 'filter':
                    if (item.text.length) {
                        filterRange
                            .style("horizontalAlignment", item.align)
                            .style("fontSize", 10)
                            .value(`${item.label}: ${item.text}`)
                            .merged(true);
                        rowCount++;
                    }
                    break;
                case 'new-line':
                    filterRange
                        .value('');
                    rowCount++;
                    break;
            }
        });

        let filterRange = workSheet.range(`A${rowCount}:${this._getColumnName(names.length - 1)}${rowCount}`);
        filterRange.style("fill", 2);

        return rowCount - 1;
    }

    /**
     * Метод, который возвращает стиль по заданному идентификатору
     * @param {Workbook} workBook объект книги Экселя
     * @param {Number} id индентификатор стиля
     * @private
     */
    _getStyleById(workBook, id) {
        return workBook.styleSheet().createStyle(id);
    }

    /**
     * Метод, который формирует секцию с фильтрами
     * @param {String} filters - строка с фильтрами
     * @param {Array<String>} names - массив наименований колонок
     * @param {*} workSheet - объект Worksheet
     * @param {Number} styleId - идентификатор стиля для секции
     * @returns {Number} количество строк, которое занимает секция
     * @private
     */
    _prepareFiltersSection(filters, names, workSheet, styleId) {
        let filterRange = workSheet.range(`A1:${this._getColumnName(names.length - 1)}${filters.split('\r\n').length}`);

        filterRange.value(filters);
        filterRange.merged(true);

        if (styleId) {
            //   let style = this._getStyleById(workSheet.workbook(), styleId);
            //   filterRange.style(style);
        }

        return filterRange._numRows;
    }

    /**
     * Метод, который формирует секцию с наименованием колонок
     * @param {Number} startRow - номер строки, с которой начинается секция
     * @param {Array<String>} headers - массив наименований колонок
     * @param {*} workSheet - объект Worksheet
     * @param {Number} styleId - идентификатор стиля для секции
     * @returns {Number} количество строк, которое занимает секция
     * @private
     */
    _prepareHeadersSection(startRow, headers, workSheet, styleId) {
        let headersRange = workSheet.range(`A${startRow}:${this._getColumnName(headers.length - 1)}${startRow}`);
        if (styleId) {
            // let style = this._getStyleById(workSheet.workbook(), styleId);
            // headersRange.style(style);
        }
        headersRange.value([headers]);
        headersRange.autoFilter();

        return startRow + headersRange._numRows;
    }

    /**
    * Метод, который формирует секцию с данными
    * @param {Number} startRow - номер строки, с которой начинается секция
    * @param {Array<String>} headers - массив наименований колонок
    * @param {Array<Array>} data - массив данных
    * @param {*} workSheet - объект Worksheet
    * @param {Number} styleId - идентификатор стиля для секции
    * @param {Boolean} isTotal - Отображать ли последнюю строчку с количеством строк
    * @private
    */
    _prepareDataSection(startRow, headers, data, workSheet, styleId, isTotal = true) {
        if (data.length > 0) {
            let dataRange = workSheet.range(`A${startRow}:${this._getColumnName(headers.length - 1)}${data.length + startRow - 1}`);

            dataRange.forEach((cell, ri, ci) => {
                cell.value(data[ri][ci]);
            });

            if (isTotal) {
                let total = data.length;
                workSheet.cell(`A${dataRange._numRows + startRow}`).value(`Всего: ${total} ${this.declOfNum(total, ['запись', 'записи', 'записей'])}`).style({ bold: true });
            }

            if (styleId) {
                let style = this._getStyleById(workSheet.workbook(), styleId);
                //dataRange.style(style);
            }
        }
    }

    /**
     * Метод, который возвращает правильное склонение слова для числительных
     * @param {Number} number - число
     * @param {Array<string>} - titles варианты склонений
     */
    declOfNum(number, titles) {
        let cases = [2, 0, 1, 1, 1, 2];
        return titles[(number % 100 > 4 && number % 100 < 20) ? 2 : cases[(number % 10 < 5) ? number % 10 : 5]];
    }

    /**
     * Метод, который возвращает наименование колонки по индексу
     * @param {Number} index - индекс колонки
     * @returns {String} - наименование колонки
     * @private
     */
    _getColumnName(index) {
        const colNames = 'abcdefghijklmnopqrstuvwxyz';
        if (index < 26) return colNames[index].toUpperCase();
        let one = colNames[Math.floor(index / 25) - 1].toUpperCase();
        let two = colNames[index % 26].toUpperCase();
        return `${one}${two}`;
    }

    /**
     * Метод, который устанавливает ширину колонки
     * @param {*} workSheet - объект WorkSheet
     * @param {Number} index - индекс колонки
     * @param {Number} width - ширина колонки
     */
    _setColumnWidth(workSheet, index, width) {
        workSheet.column(this._getColumnName(index)).width(width / 6);
    }
}

(async function () {
    let script = document.createElement('script');
    script.src = "xlsx-populate/browser/xlsx-populate.min.js";
    script.async = true;
    script.onload = (() => {
        customElements.define('pl-data-exporter', PlDataExporter);
    })

    document.head.appendChild(script);
}());
