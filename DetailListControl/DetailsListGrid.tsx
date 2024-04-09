import * as React from 'react';
import { Overlay } from '@fluentui/react/lib/Overlay';
import {
  ScrollablePane,
  ScrollbarVisibility
} from '@fluentui/react/lib/ScrollablePane';
import { Stack } from '@fluentui/react/lib/Stack';
import { Sticky } from '@fluentui/react/lib/Sticky';
import { StickyPositionType } from '@fluentui/react/lib/Sticky';
import { IObjectWithKey } from '@fluentui/react/lib/Selection';
import { IRenderFunction } from '@fluentui/react/lib/Utilities';
import { useConst, useForceUpdate } from "@fluentui/react-hooks";
import { Selection } from "@fluentui/react/lib/Selection";
import { SelectionMode } from "@fluentui/react/lib/Utilities";
import { ContextualMenu, DirectionalHint, IContextualMenuProps } from '@fluentui/react/lib/ContextualMenu';
import { IconButton } from '@fluentui/react/lib/Button';
import { Link } from '@fluentui/react/lib/Link';
import {
  DetailsList,
  ConstrainMode,
  DetailsListLayoutMode,
  IColumn,
  IDetailsHeaderProps,
  IDetailsListProps,
  IDetailsRowStyles,
  DetailsRow
} from '@fluentui/react/lib/DetailsList';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Label } from '@fluentui/react/lib/Label';

type DataSet = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;

export interface GridProps {
  width?: number;
  height?: number;
  columns: ComponentFramework.PropertyHelper.DataSetApi.Column[];
  records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
  // totalRecords:number;
  sortedRecordIds: string[];
  hasNextPage: boolean;
  hasPreviousPage: boolean;
  totalResultCount: number;
  currentPage: number;
  sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[];
  filtering: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;
  resources: ComponentFramework.Resources;
  itemsLoading: boolean;
  // highlightValue: string | null;
  // highlightColor: string | null;
  // DropdownField: string | null;
  setSelectedRecords: (ids: string[]) => void;
  onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
  onSort: (name: string, desc: boolean) => void;
  onFilter: (name: string, filtered: boolean) => void;
  loadFirstPage: () => void;
  loadNextPage: () => void;
  loadPreviousPage: () => void;
  onFullScreen: () => void;
  isFullScreen: boolean;
}
const options: IDropdownOption[] = [
  
  { key: 'Misallocated Charge', text: 'Misallocated Charge' },
  { key: 'Onboarding', text: 'Onboarding' },
  { key: 'Recruitment', text: 'Recruitment'},
  { key: 'Role Enablement', text: 'Role Enablement' },
  { key: 'Team Development', text: 'Team Development' },
  { key: 'Team Engagement', text: 'Engagement' },
  // { key: 'lettuce', text: 'Lettuce' },
];
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 200 },
};
const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
  if (props && defaultRender) {
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender({
          ...props,
        })}
      </Sticky>
    );
  }
  return null;
};

const onRenderItemColumn = (
  item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
  index?: number,
  column?: IColumn,

) => {


  if (column && column.fieldName && item) {
    if (column.fieldName=="cr235_dropdown" || column.fieldName=="Dropdown" ||column.name=="Dropdown") {
      return (
        <Dropdown
          placeholder="Select activity driver"
          options={options}
          styles={dropdownStyles}
          
        />
      )
    }
    return <>{item?.getFormattedValue(column.fieldName)}</>;

  }
  return <></>;
};

export const Grid = React.memo((props: GridProps) => {
  const {
    records,
    sortedRecordIds,
    columns,
    width,
    height,
    hasNextPage,
    hasPreviousPage,
    sorting,
    filtering,
    currentPage,
    itemsLoading,
    setSelectedRecords,
    onNavigate,
    onSort,
    onFilter,
    resources,
    loadFirstPage,
    loadNextPage,
    loadPreviousPage,
    onFullScreen,
    isFullScreen,
    // highlightValue,
    // highlightColor,
    // totalRecords,
    // DropdownField,
    totalResultCount
  } = props;

  const [isComponentLoading, setIsLoading] = React.useState<boolean>(false);

  const [contextualMenuProps, setContextualMenuProps] =
    React.useState<IContextualMenuProps>();

  const onContextualMenuDismissed = React.useCallback(() => {
    setContextualMenuProps(undefined);
  }, [setContextualMenuProps]);

  const getContextualMenuProps = React.useCallback(
    (
      column: IColumn,
      ev: React.MouseEvent<HTMLElement>
    ): IContextualMenuProps => {
      const menuItems = [
        {
          key: "aToZ",
          name: resources.getString("Label_SortAZ"),
          iconProps: { iconName: "SortUp" },
          canCheck: true,
          checked: column.isSorted && !column.isSortedDescending,
          disable: (
            column.data as ComponentFramework.PropertyHelper.DataSetApi.Column
          ).disableSorting,
          onClick: () => {
            onSort(column.key, false);
            setContextualMenuProps(undefined);
            setIsLoading(true);
          },
        },
        {
          key: "zToA",
          name: resources.getString("Label_SortZA"),
          iconProps: { iconName: "SortDown" },
          canCheck: true,
          checked: column.isSorted && column.isSortedDescending,
          disable: (
            column.data as ComponentFramework.PropertyHelper.DataSetApi.Column
          ).disableSorting,
          onClick: () => {
            onSort(column.key, true);
            setContextualMenuProps(undefined);
            setIsLoading(true);
          },
        },
        {
          key: "filter",
          name: resources.getString("Label_DoesNotContainData"),
          iconProps: { iconName: "Filter" },
          canCheck: true,
          checked: column.isFiltered,
          onClick: () => {
            onFilter(column.key, column.isFiltered !== true);
            setContextualMenuProps(undefined);
            setIsLoading(true);
          },
        },
      ];
      return {
        items: menuItems,
        target: ev.currentTarget as HTMLElement,
        directionalHint: DirectionalHint.bottomLeftEdge,
        gapSpace: 10,
        isBeakVisible: true,
        onDismiss: onContextualMenuDismissed,
      };
    },
    [setIsLoading, onFilter, setContextualMenuProps]
  );

  const onColumnContextMenu = React.useCallback(
    (column?: IColumn, ev?: React.MouseEvent<HTMLElement>) => {
      if (column && ev) {
        setContextualMenuProps(getContextualMenuProps(column, ev));
      }
    },
    [getContextualMenuProps, setContextualMenuProps]
  );

  const onColumnClick = React.useCallback(
    (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
      if (column && ev) {
        setContextualMenuProps(getContextualMenuProps(column, ev));
      }
    },
    [getContextualMenuProps, setContextualMenuProps]
  );

  const items: (DataSet | undefined)[] = React.useMemo(() => {
    setIsLoading(false);

    const sortedRecords: (DataSet | undefined)[] = sortedRecordIds.map((id) => {
      const record = records[id];
      return record;
    });

    return sortedRecords;
  }, [records, sortedRecordIds, hasNextPage, setIsLoading]);

  const gridColumns = React.useMemo(() => {
    return columns
      .filter((col) => !col.isHidden && col.order >= 0)
      .sort((a, b) => a.order - b.order)
      .map((col) => {
        const sortOn = sorting && sorting.find((s) => s.name === col.name);
        const filtered =
          filtering &&
          filtering.conditions &&
          filtering.conditions.find((f) => f.attributeName == col.name);
        return {
          key: col.name,
          name: col.displayName,
          fieldName: col.name,
          isSorted: sortOn != null,
          isSortedDescending: sortOn?.sortDirection === 1,
          isResizable: true,
          isFiltered: filtered != null,
          data: col,
          minWidth: 200,
          onColumnContextMenu: onColumnContextMenu,
          onColumnClick: onColumnClick,
        } as IColumn;
      });
  }, [columns, sorting, onColumnContextMenu, onColumnClick]);

  const rootContainerStyle: React.CSSProperties = React.useMemo(() => {
    return {
      height: height,
      width: width,
    };
  }, [width, height]);

  const onRenderRow: IDetailsListProps['onRenderRow'] = (props) => {
    const customStyles: Partial<IDetailsRowStyles> = {
    };
    if (props && props.item) {

      const item = props.item as DataSet | undefined;
      // if (highlightColor && highlightValue && item?.getValue('HighlightIndicator') == highlightValue) {
      //   customStyles.root = { backgroundColor: highlightColor };
      // }
      return <DetailsRow {...props} styles={customStyles}  />;
    }
    return null;
  };
  const forceUpdate = useForceUpdate();
  const onSelectionChanged = React.useCallback(() => {
    const items = selection.getItems() as DataSet[];
    const selected = selection.getSelectedIndices().map((index: number) => {
      const item: DataSet | undefined = items[index];
      return item && items[index].getRecordId();
    });

    setSelectedRecords(selected);
    forceUpdate();
  }, [forceUpdate]);

  const selection: Selection = useConst(() => {
    return new Selection({
      selectionMode: SelectionMode.single,
      onSelectionChanged: onSelectionChanged,
    });
  });
  function stringFormat(template: string, ...args: string[]): string {
    for (const k in args) {
      template = template.replace("{" + k + "}", args[k]);
    }
    return template;
  }
  return (
    <Stack verticalFill grow style={rootContainerStyle}>
      <Stack.Item grow style={{ position: 'relative', backgroundColor: 'white' }}>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <DetailsList
            columns={gridColumns}
            onRenderItemColumn={onRenderItemColumn}
            onRenderDetailsHeader={onRenderDetailsHeader}
            items={items}
            setKey={`set${currentPage}`} // Ensures that the selection is reset when paging
            initialFocusedIndex={0}
            checkButtonAriaLabel="select row"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            constrainMode={ConstrainMode.unconstrained}
            selection={selection}
            onItemInvoked={onNavigate}
            onRenderRow={onRenderRow}
          ></DetailsList>
          {contextualMenuProps && <ContextualMenu {...contextualMenuProps} />}
        </ScrollablePane>
        {(itemsLoading || isComponentLoading) && <Overlay />}
      </Stack.Item>
      <Stack.Item>
        <Stack horizontal style={{ width: '100%', paddingLeft: 8, paddingRight: 8 }}>
          <Stack.Item grow align="center">
            {!isFullScreen && (
              <Link onClick={onFullScreen}>{resources.getString('Label_ShowFullScreen')}</Link>
            )}
          </Stack.Item>
          <Stack.Item >
          <Label> {totalResultCount +" records "}</Label>
         
          </Stack.Item>
          <IconButton
            alt="First Page"
            iconProps={{ iconName: 'Rewind' }}
            disabled={!hasPreviousPage}
            onClick={loadFirstPage}
          />
          <IconButton
            alt="Previous Page"
            iconProps={{ iconName: 'Previous' }}
            disabled={!hasPreviousPage}
            onClick={loadPreviousPage}
          />
          <Stack.Item align="center">
            {stringFormat(
              resources.getString('Label_Grid_Footer'),
              currentPage.toString(),
              selection.getSelectedCount().toString(),
            )}
          </Stack.Item>
          <IconButton
            alt="Next Page"
            iconProps={{ iconName: 'Next' }}
            disabled={!hasNextPage}
            onClick={loadNextPage}
          />
        </Stack>
      </Stack.Item>
    </Stack>
  );
});

Grid.displayName = 'Grid';