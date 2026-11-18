import * as React from 'react';
import * as AdaptiveCards from 'adaptivecards';

export interface IAdaptiveCardHostProps {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  payload: any;
  onAction?: (action: AdaptiveCards.Action) => void;
}

/**
 * AdaptiveCardHost
 * Renders an Adaptive Card payload inside the chat bubble.
 */
export const AdaptiveCardHost: React.FC<IAdaptiveCardHostProps> = ({ payload, onAction }) => {
  const containerRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (!containerRef.current || !payload) return;

    const card = new AdaptiveCards.AdaptiveCard();

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

    const rendered = card.render();
    if (rendered) {
      containerRef.current.innerHTML = '';
      containerRef.current.appendChild(rendered);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [payload]);

  return (
    <div
      ref={containerRef}
      className="ce-adaptive-card"
      style={{ width: '100%' }}
    />
  );
};