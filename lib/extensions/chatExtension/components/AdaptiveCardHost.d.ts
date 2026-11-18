import * as React from 'react';
import * as AdaptiveCards from 'adaptivecards';
export interface IAdaptiveCardHostProps {
    payload: any;
    onAction?: (action: AdaptiveCards.Action) => void;
}
/**
 * AdaptiveCardHost
 * Renders an Adaptive Card payload inside the chat bubble.
 */
export declare const AdaptiveCardHost: React.FC<IAdaptiveCardHostProps>;
//# sourceMappingURL=AdaptiveCardHost.d.ts.map