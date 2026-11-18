import * as React from 'react';
import * as AdaptiveCards from 'adaptivecards';
/**
 * AdaptiveCardHost
 * Renders an Adaptive Card payload inside the chat bubble.
 */
export var AdaptiveCardHost = function (_a) {
    var payload = _a.payload, onAction = _a.onAction;
    var containerRef = React.useRef(null);
    React.useEffect(function () {
        if (!containerRef.current || !payload)
            return;
        var card = new AdaptiveCards.AdaptiveCard();
        // Apply a minimal host config matching our dark-teal theme
        card.hostConfig = new AdaptiveCards.HostConfig({
            fontFamily: "'Inter Variable', 'Segoe UI', sans-serif",
            containerStyles: {
                default: {
                    backgroundColor: 'transparent',
                    foregroundColors: {
                        default: { default: '#e2e8f0', subtle: '#94a3b8' },
                    },
                },
            },
            actions: {
                actionAlignment: 'stretch',
                buttonSpacing: 8,
                showCard: { actionMode: 'inline' },
            },
        });
        card.parse(payload);
        if (onAction) {
            card.onExecuteAction = onAction;
        }
        var rendered = card.render();
        if (rendered) {
            containerRef.current.innerHTML = '';
            containerRef.current.appendChild(rendered);
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [payload]);
    return (React.createElement("div", { ref: containerRef, className: "ce-adaptive-card", style: { width: '100%' } }));
};
//# sourceMappingURL=AdaptiveCardHost.js.map