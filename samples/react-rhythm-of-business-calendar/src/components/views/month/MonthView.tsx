import React, { FC, useCallback, useRef } from 'react';
import { EventOccurrence, ViewKeys } from 'model';
import { EventDetailsCallout, IEventDetailsCallout } from '../../events';
import { IViewDescriptor } from '../IViewDescriptor';
import { IViewProps } from '../IViewProps';
import { Builder } from './Builder';
import { Header } from './Header';
import { Week } from './Week';

import { ViewNames as strings } from 'ComponentStrings';
import { FocusZone } from '@fluentui/react';
import { useTimeZoneService } from "services";

const MonthView: FC<IViewProps> = ({ anchorDate, eventCommands, viewCommands, cccurrences, selectedTemplateKeys }) => {
    const { siteTimeZone } = useTimeZoneService();
    anchorDate = anchorDate.tz(siteTimeZone.momentId,true);
    const weeks = Builder.build(cccurrences, anchorDate); 
    //console.log("weeks", weeks);
    const detailsCallout = useRef<IEventDetailsCallout>();

    const onActivate = useCallback((cccurrence: EventOccurrence, target: HTMLElement) => {
        detailsCallout.current?.open(cccurrence, target);
    }, []);

    return (
        <FocusZone>
            <Header />
            {weeks.map(week =>
                <Week
                    key={week.start.format('L')}
                    week={week}
                    anchorDate={anchorDate}
                    onActivate={onActivate}
                    viewCommands={viewCommands}
                    selectedTemplateKeys={selectedTemplateKeys}
                />
            )}
            <EventDetailsCallout
                commands={eventCommands}
                componentRef={detailsCallout}
               // channels ={channels}
            />
        </FocusZone>
    );
};

export const MonthViewDescriptor: IViewDescriptor = {
    id: ViewKeys.monthly,
    title: strings.Month,
    renderer: MonthView,
    dateRotatorController: {
        previousIconProps: { iconName: 'ChevronUp' },
        nextIconProps: { iconName: 'ChevronDown' },
        previousDate: date => date.clone().subtract(1, 'month'),
        nextDate: date => date.clone().add(1, 'month'),
        dateString: date => date.format('MMMM YYYY')
    },
    dateRange: Builder.dateRange
};