import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Grid } from './DetailsListGrid';
import { initializeIcons } from "@fluentui/react/lib/Icons";
initializeIcons(undefined, { disableWarnings: true });
type DataSet = ComponentFramework.PropertyTypes.DataSet;


export class DetailListControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    /**
     * Empty constructor.
     */
    notifyOutputChanged: () => void;
    container: HTMLDivElement;
    context: ComponentFramework.Context<IInputs>;
    sortedRecordsIds: string[] = [];
    resources: ComponentFramework.Resources;
    isTestHarness: boolean;
    records: {
        [id: string]: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
    };
    currentPage = 1;
    filteredRecordCount?: number;
    isFullScreen = false;

    setSelectedRecords = (ids: string[]): void => {
        this.context.parameters.DelegationTableTesting.setSelectedRecordIds(ids);
    };

    onNavigate = (
        item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord
    ): void => {
        if (item) {
            this.context.parameters.DelegationTableTesting.openDatasetItem(item.getNamedReference());
        }
    };
    onSort = (name: string, desc: boolean): void => {
        const sorting = this.context.parameters.DelegationTableTesting.sorting;
        while (sorting.length > 0) {
            sorting.pop();
        }
        this.context.parameters.DelegationTableTesting.sorting.push({
            name: name,
            sortDirection: desc ? 1 : 0,
        });
        this.context.parameters.DelegationTableTesting.refresh();
    };

    onFilter = (name: string, filter: boolean): void => {
        const filtering = this.context.parameters.DelegationTableTesting.filtering;
        if (filter) {
            filtering.setFilter({
                conditions: [
                    {
                        attributeName: name,
                        conditionOperator: 12, // Does not contain Data
                    },
                ],
            } as ComponentFramework.PropertyHelper.DataSetApi.FilterExpression);
        } else {
            filtering.clearFilter();
        }
        this.context.parameters.DelegationTableTesting.refresh();
    };
    loadFirstPage = (): void => {
        this.currentPage = 1;
        this.context.parameters.DelegationTableTesting.paging.loadExactPage(1);
    };
    loadNextPage = (): void => {
        this.currentPage++;
        this.context.parameters.DelegationTableTesting.paging.loadExactPage(this.currentPage);
    };
    loadPreviousPage = (): void => {
        this.currentPage--;
        this.context.parameters.DelegationTableTesting.paging.loadExactPage(this.currentPage);
    };
    onFullScreen = (): void => {
        this.context.mode.setFullScreen(true);
    };
    AlreadyPaid = (flocId: string, optionSetMonthValue: string): Promise<any[] | null> => {
        return new Promise((resolve, reject) => {
            const numverFinal = parseInt(optionSetMonthValue);
            const fetchXml = `?fetchXml=<fetch version='1.0' mapping='logical' no-lock='false' distinct='true'><entity name='cr235_DelegationTableTesting'><attribute name='cr235_SubID'/></entity></fetch>`;

            Xrm.WebApi.retrieveMultipleRecords("rtlme_cashflow", fetchXml)
                .then((result: Xrm.RetrieveMultipleResult) => {
                    const entities=result.entities
                    resolve(entities);
                })
                .catch((err: Xrm.ErrorResponse) => {
                    console.error(err);
                    reject(err);
                });
        });
    };


    constructor() {

    }


    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        // Add control initialization code
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        this.context = context;
        this.context.mode.trackContainerResize(true);
        this.resources = this.context.resources;
        this.isTestHarness = document.getElementById("control-dimensions") !== null;
        var i=this.AlreadyPaid("1","1")
        console.log(i)
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
        const dataset = context.parameters.DelegationTableTesting;
        const paging = context.parameters.DelegationTableTesting.paging;
        const datasetChanged = context.updatedProperties.indexOf("dataset") > -1;
        const resetPaging =
            datasetChanged &&
            !dataset.loading &&
            !dataset.paging.hasPreviousPage &&
            this.currentPage !== 1;

        if (resetPaging) {
            this.currentPage = 1;
        }
        if (context.updatedProperties.indexOf('fullscreen_close') > -1) {
            this.isFullScreen = false;
        }
        if (context.updatedProperties.indexOf('fullscreen_open') > -1) {
            this.isFullScreen = true;
        }
        if (resetPaging || datasetChanged || this.isTestHarness) {
            this.records = dataset.records;
            this.sortedRecordsIds = dataset.sortedRecordIds;
        }

        // The test harness provides width/height as strings
        const allocatedWidth = parseInt(
            context.mode.allocatedWidth as unknown as string
        );
        const allocatedHeight = parseInt(
            context.mode.allocatedHeight as unknown as string
        );

        if (this.filteredRecordCount !== this.sortedRecordsIds.length) {
            this.filteredRecordCount = this.sortedRecordsIds.length;
            this.notifyOutputChanged();
        }

        ReactDOM.render(
            React.createElement(Grid, {
                width: allocatedWidth,
                height: allocatedHeight,
                columns: dataset.columns,
                records: this.records,
                // totalRecords:paging.totalResultCount,
                sortedRecordIds: this.sortedRecordsIds,
                hasNextPage: paging.hasNextPage,
                hasPreviousPage: paging.hasPreviousPage,
                currentPage: this.currentPage,
                totalResultCount: paging.totalResultCount,
                sorting: dataset.sorting,
                filtering: dataset.filtering && dataset.filtering.getFilter(),
                resources: this.resources,
                itemsLoading: dataset.loading,
                // highlightValue: this.context.parameters.HighlightValue.raw,
                // highlightColor: this.context.parameters.HighlightColor.raw,
                // DropdownField:context.parameters.DropdownField.raw,
                setSelectedRecords: this.setSelectedRecords,
                onNavigate: this.onNavigate,
                onSort: this.onSort,
                onFilter: this.onFilter,
                loadFirstPage: this.loadFirstPage,
                loadNextPage: this.loadNextPage,
                loadPreviousPage: this.loadPreviousPage,
                isFullScreen: this.isFullScreen,
                onFullScreen: this.onFullScreen,
            }),
            this.container
        );



    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {
            FilteredRecordCount: this.filteredRecordCount,
        } as IOutputs;
    }
    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
        ReactDOM.unmountComponentAtNode(this.container);

    }
}
