import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  BotSparkle24Filled,
  ArrowCounterclockwise20Regular,
  ArrowMaximize20Regular,
  ArrowMinimize20Regular,
  Dismiss20Regular,
  Send20Filled,
  LockClosed24Regular,
  Settings20Regular,
  Person20Regular,
  ShieldLock20Regular,
  ClipboardTaskListLtr20Regular,
  Wrench20Regular,
  DocumentBulletList20Regular,
  ChatMultiple20Regular,
  Sparkle20Regular,
  ThumbLike20Regular,
  ThumbDislike20Regular,
  Chat20Regular,
  Alert20Regular,
  Map20Regular,
  Warning20Regular,
  Location20Regular,
  Tag20Regular,
  WeatherMoon20Regular,
  MusicNote120Regular,
  Food20Regular,
  Building20Regular,
} from '@fluentui/react-icons';
import { IChatMessage, ISuggestedAction } from '../models/IChatMessage';
import { IChatConfig }  from '../models/IChatConfig';
import { ILiveUpdate }         from '../models/ILiveUpdate';
import { IRecommendation }     from '../models/IRecommendation';
import { ChatProxyService }    from '../services/ChatProxyservices';
import { LiveUpdatesService }  from '../services/Liveupdatesservice';
import { RecommendationsService } from '../services/Recommendationsservice';
import { AdaptiveCardHost } from './AdaptiveCardHost';
import styles from './ChatWidget.module.scss';

// ─── Constants ────────────────────────────────────────────────────────────────
const SHOW_ONCE_KEY = 'ce_label_shown';

type ActiveTab = 'chat' | 'updates' | 'nearby';
type FeedbackStep = 'idle' | 'rating' | 'forms' | 'done';

function formatTime(d: Date): string {
  return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

function stripUsernameFromWelcome(text: string): string {
  if (!text) return text;
  return text.replace(
    /^(Hello|Hi|Hey|Greetings|Welcome)(?:,?\s+)([A-Z][a-z]+)([,!.]?)/,
    '$1$3'
  );
}

function sanitizeHttpUrl(raw?: string): string | null {
  if (!raw) return null;
  const t = raw.trim();
  return /^https?:\/\//i.test(t) ? t : null;
}

/* ── Citations helpers (unchanged) ─────────────────────────────────────────── */
interface ICitation { index: number; url: string; title?: string; }

function niceLabelFromUrl(u: string, explicitTitle?: string): string {
  if (explicitTitle) return explicitTitle;
  try {
    const url = new URL(u);
    const parts = url.pathname.split('/').filter(function(s) { return s.length > 0; });
    const last = parts.length > 0 ? parts[parts.length - 1] : url.hostname;
    const base = decodeURIComponent(last).replace(/\.[a-z0-9]+$/i, '');
    const label = base || url.hostname;
    return label.length > 36 ? label.slice(0, 33) + '…' : label;
  } catch {
    return u.length > 36 ? u.slice(0, 33) + '…' : u;
  }
}

function extractCitationsFromText(rawText: string): { cleanText: string; citations: ICitation[] } {
  const lines = (rawText || '').split('\n');
  const body: string[] = [];
  const citObj: Record<number, ICitation> = {};
  const rBracketSame  = /^\s*\[(\d+)\]\s*:\s*(https?:\S+)(?:\s+"([^"]+)")?\s*$/i;
  const rBracketAlone = /^\s*\[(\d+)\]\s*:\s*$/i;
  const rNumberDot    = /^\s*(\d+)[.)]\s+(https?:\S+)(?:\s+"([^"]+)")?\s*$/i;
  const rUrlOnly      = /^\s*(https?:\S+)(?:\s+"([^"]+)")?\s*$/i;
  const rRefsHeading  = /^\s*#{0,3}\s*references?:?\s*$/i;
  let i = 0; let inRefsBlock = false; let maxIdx = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (rRefsHeading.test(line)) { inRefsBlock = true; i++; continue; }
    const mSame = rBracketSame.exec(line);
    if (mSame) { const idx=parseInt(mSame[1],10); const url=sanitizeHttpUrl(mSame[2]); if(url){citObj[idx]={index:idx,url:url,title:mSame[3]};if(idx>maxIdx)maxIdx=idx;} inRefsBlock=true; i++; continue; }
    const mAlone = rBracketAlone.exec(line);
    if (mAlone) { const idx=parseInt(mAlone[1],10); let j=i+1; while(j<lines.length&&lines[j].trim()==='')j++; if(j<lines.length){const mNext=rUrlOnly.exec(lines[j]);if(mNext){const url=sanitizeHttpUrl(mNext[1]);if(url){citObj[idx]={index:idx,url:url,title:mNext[2]};if(idx>maxIdx)maxIdx=idx;}i=j+1;inRefsBlock=true;continue;}} inRefsBlock=true; i++; continue; }
    const mNum = rNumberDot.exec(line);
    if (mNum) { const idx=parseInt(mNum[1],10); const url=sanitizeHttpUrl(mNum[2]); if(url){citObj[idx]={index:idx,url:url,title:mNum[3]};if(idx>maxIdx)maxIdx=idx;} inRefsBlock=true; i++; continue; }
    if (inRefsBlock) { const mUrl=rUrlOnly.exec(line); if(mUrl){const ni=maxIdx+1;const url=sanitizeHttpUrl(mUrl[1]);if(url){citObj[ni]={index:ni,url:url,title:mUrl[2]};maxIdx=ni;}i++;continue;} }
    body.push(line); i++;
  }
  const citations: ICitation[] = [];
  const keys = Object.keys(citObj);
  for (let k = 0; k < keys.length; k++) { citations.push(citObj[parseInt(keys[k], 10)]); }
  citations.sort(function(a, b) { return a.index - b.index; });
  return { cleanText: body.join('\n').trim(), citations: citations };
}

/* ── Markdown renderer ─────────────────────────────────────────────────────── */
interface IRenderOpts { citationsByIndex?: Record<number,ICitation>; onCitationClick?: (n:number)=>void; }

function renderMarkdown(text: string, opts?: IRenderOpts): React.ReactNode {
  if (!text) { return null; }
  const lines = text.split('\n');
  const nodes: React.ReactNode[] = [];
  let listItems: React.ReactNode[] = [];
  let listType: 'ul'|'ol'|null = null;
  let paraLines: string[] = [];
  let key = 0;
  function flushList(): void { if(!listItems.length)return; if(listType==='ol'){nodes.push(<ol key={'ol-'+key++} className={styles.mdOl}>{listItems}</ol>);}else{nodes.push(<ul key={'ul-'+key++} className={styles.mdUl}>{listItems}</ul>);} listItems=[];listType=null; }
  function flushPara(): void { if(!paraLines.length)return; nodes.push(<p key={'p-'+key++} className={styles.mdP}>{inlineFormat(paraLines.join(' '))}</p>); paraLines=[]; }

  function inlineFormat(raw: string): React.ReactNode {
    const parts: React.ReactNode[] = []; let remaining = raw; let idx = 0;
    while (remaining.length > 0) {
      const boldM = remaining.match(/^(.*?)\*\*(.+?)\*\*(.*)/);
      const italM = remaining.match(/^(.*?)\*(.+?)\*(.*)/);
      const codeM = remaining.match(/^(.*?)`(.+?)`(.*)/);
      const linkM = remaining.match(/^(.*?)\[(.+?)\]\((.+?)\)(.*)/);
      const citeM = remaining.match(/^(.*?)\[(\d+)\](.*)/);

      // ── NEW: bare URL and autolink patterns ────────────────────────────────
      // autolink:  <https://example.com>     → linkify, render without brackets
      // bare url:  https://example.com       → linkify in place
      // Stops on whitespace and chars that almost always end a URL in prose.
      // FIX: added \] and \[ to the excluded class below so a bare URL stops
      // cleanly before a following [1] citation marker instead of consuming
      // the bracket and digit into the href/link text (e.g. "...page[1] [2]."
      // was previously linking "...page[1" and breaking the first citation).
      const autolinkM = remaining.match(/^(.*?)<(https?:\/\/[^\s>]+)>(.*)/);
      const urlM      = remaining.match(/^(.*?)(https?:\/\/[^\s<>"')\]\[]+)(.*)/);

      const matches: Array<{pos:number;type:string;m:RegExpMatchArray}> = [];
      if(boldM&&boldM[1]!==undefined)matches.push({pos:boldM[1].length,type:'bold',m:boldM});
      if(italM&&italM[1]!==undefined)matches.push({pos:italM[1].length,type:'italic',m:italM});
      if(codeM&&codeM[1]!==undefined)matches.push({pos:codeM[1].length,type:'code',m:codeM});
      if(linkM&&linkM[1]!==undefined)matches.push({pos:linkM[1].length,type:'link',m:linkM});
      if(citeM&&citeM[1]!==undefined)matches.push({pos:citeM[1].length,type:'cite',m:citeM});
      if(autolinkM&&autolinkM[1]!==undefined)matches.push({pos:autolinkM[1].length,type:'autolink',m:autolinkM});
      if(urlM&&urlM[1]!==undefined)matches.push({pos:urlM[1].length,type:'url',m:urlM});

      if(!matches.length){parts.push(remaining);break;}
      matches.sort(function(a,b){return a.pos-b.pos;});
      const best=matches[0];
      if(best.m[1])parts.push(best.m[1]);
      if(best.type==='bold'){parts.push(<strong key={idx++}>{best.m[2]}</strong>);remaining=best.m[3];}
      else if(best.type==='italic'){parts.push(<em key={idx++}>{best.m[2]}</em>);remaining=best.m[3];}
      else if(best.type==='code'){parts.push(<code key={idx++} className={styles.mdCode}>{best.m[2]}</code>);remaining=best.m[3];}
      else if(best.type==='link'){const raw2=best.m[3];const safe=(raw2.indexOf('http://')===0||raw2.indexOf('https://')===0)?raw2:'#';parts.push(<a key={idx++} href={safe} target="_blank" rel="noopener noreferrer" className={styles.mdLink}>{best.m[2]}</a>);remaining=best.m[4]||'';}
      else if(best.type==='cite'){const n=parseInt(best.m[2],10);const cit=(opts&&opts.citationsByIndex)?opts.citationsByIndex[n]:undefined;if(cit&&cit.url){const safe=sanitizeHttpUrl(cit.url)||'#';const ct=cit.title?'Open '+cit.title:'Open reference '+String(n);const oc=function(){if(opts&&opts.onCitationClick)opts.onCitationClick(n);};parts.push(<a key={idx++} href={safe} target="_blank" rel="noopener noreferrer" className={styles.citeSup} onClick={oc} title={ct}>{n}</a>);}else{parts.push('['+best.m[2]+']');}remaining=best.m[3]||'';}
      // ── NEW: autolink branch ────────────────────────────────────────────────
      else if(best.type==='autolink'){
        const url = best.m[2];
        const safe = sanitizeHttpUrl(url) || '#';
        parts.push(
          <a key={idx++} href={safe} target="_blank" rel="noopener noreferrer" className={styles.mdLink}>
            {url}
          </a>
        );
        remaining = best.m[3] || '';
      }
      // ── NEW: bare URL branch ────────────────────────────────────────────────
      else if(best.type==='url'){
        let url: string = best.m[2];
        let tail: string = '';
        // Strip trailing punctuation that almost never belongs to a URL,
        // e.g. "see https://example.com." or "(https://x.com)"
        while (url.length > 0 && /[.,;:!?)\]}'"`]/.test(url.charAt(url.length - 1))) {
          tail = url.charAt(url.length - 1) + tail;
          url = url.slice(0, -1);
        }
        const safe = sanitizeHttpUrl(url) || '#';
        parts.push(
          <a key={idx++} href={safe} target="_blank" rel="noopener noreferrer" className={styles.mdLink}>
            {url}
          </a>
        );
        remaining = tail + (best.m[3] || '');
      }
    }
    return parts.length===1?parts[0]:<React.Fragment>{parts}</React.Fragment>;
  }

  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if(t===''){flushList();flushPara();continue;}
    const hM=t.match(/^(#{1,3})\s+(.+)/);
    if(hM){flushList();flushPara();const lvl=hM[1].length;const hc=inlineFormat(hM[2]);if(lvl===1)nodes.push(<div key={'h-'+key++} className={styles.mdH1}>{hc}</div>);else if(lvl===2)nodes.push(<div key={'h-'+key++} className={styles.mdH2}>{hc}</div>);else nodes.push(<div key={'h-'+key++} className={styles.mdH3}>{hc}</div>);continue;}
    const ulM=t.match(/^[-*]\s+(.+)/);
    if(ulM){flushPara();if(listType==='ol')flushList();listType='ul';listItems.push(<li key={'li-'+key++} className={styles.mdLi}>{inlineFormat(ulM[1])}</li>);continue;}
    const olM=t.match(/^\d+\.\s+(.+)/);
    if(olM){flushPara();if(listType==='ul')flushList();listType='ol';listItems.push(<li key={'li-'+key++} className={styles.mdLi}>{inlineFormat(olM[1])}</li>);continue;}
    flushList(); paraLines.push(t);
  }
  flushList(); flushPara();
  return <React.Fragment>{nodes}</React.Fragment>;
}

/* ── Prompt icon map (unchanged) ────────────────────────────────────────────── */
function getPromptIcon(iconName: string): React.ReactNode {
  const n=(iconName||'').toLowerCase();
  if(n.indexOf('process')!==-1||n.indexOf('gear')!==-1||n.indexOf('setting')!==-1||n.indexOf('workflow')!==-1)return <Settings20Regular/>;
  if(n.indexOf('contact')!==-1||n.indexOf('people')!==-1||n.indexOf('poi')!==-1||n.indexOf('local')!==-1)return <Person20Regular/>;
  if(n.indexOf('access')!==-1||n.indexOf('permission')!==-1||n.indexOf('security')!==-1||n.indexOf('lock')!==-1||n.indexOf('password')!==-1)return <ShieldLock20Regular/>;
  if(n.indexOf('onboard')!==-1||n.indexOf('checklist')!==-1||n.indexOf('task')!==-1||n.indexOf('setup')!==-1)return <ClipboardTaskListLtr20Regular/>;
  if(n.indexOf('troubl')!==-1||n.indexOf('wrench')!==-1||n.indexOf('tool')!==-1||n.indexOf('fix')!==-1||n.indexOf('help')!==-1)return <Wrench20Regular/>;
  if(n.indexOf('polic')!==-1||n.indexOf('document')!==-1||n.indexOf('rule')!==-1||n.indexOf('guidanc')!==-1)return <DocumentBulletList20Regular/>;
  return <ChatMultiple20Regular/>;
}

/* ── Category icon for recommendations (unchanged) ──────────────────────────── */
function getCategoryIcon(category: string): React.ReactNode {
  const c = (category || '').toLowerCase();
  if (c.indexOf('jazz') !== -1 || c.indexOf('music') !== -1) { return <MusicNote120Regular />; }
  if (c.indexOf('food') !== -1 || c.indexOf('cafe') !== -1 || c.indexOf('restaurant') !== -1) { return <Food20Regular />; }
  if (c.indexOf('sight') !== -1 || c.indexOf('tour') !== -1 || c.indexOf('walk') !== -1) { return <Building20Regular />; }
  return <Location20Regular />;
}

/* ── BotAvatarIcon ─────────────────────────────────────────────────────────────
   Renders the tenant-configured agent icon (config.agentIconUrl) when present
   and valid; falls back to the default BotSparkle24Filled glyph if the field
   is blank, not an http(s) URL, or the image fails to load at runtime.
   NOTE: requires an `agentIconUrl?: string` field on IChatConfig — see
   ../models/IChatConfig.ts. If that field isn't added yet, pass undefined
   and this component silently renders the default icon everywhere.        ── */
interface IBotAvatarIconProps { iconUrl?: string; className: string; }
const BotAvatarIcon: React.FC<IBotAvatarIconProps> = function(props) {
  const [failed, setFailed] = React.useState(false);
  const safeUrl = sanitizeHttpUrl(props.iconUrl);
  if (safeUrl && !failed) {
    return (
      <img
        src={safeUrl}
        alt=""
        className={props.className}
        onError={function() { setFailed(true); }}
      />
    );
  }
  return <BotSparkle24Filled className={props.className} />;
};

/* ── Interfaces ─────────────────────────────────────────────────────────────── */
interface ISuggestedPrompt { title: string; subtitle: string; icon?: string; prompt: string; }

interface IChatWidgetProps {
  spHttpClient:      SPHttpClient;
  hubSiteUrl:        string;
  listTitle:         string;
  defaultTitleMatch: string;
  siteId:            string;
  siteUrl:           string;
  userRole?:         string;
}

/* ── FeedbackDialog (unchanged) ─────────────────────────────────────────────── */
interface IFeedbackDialogProps { step:FeedbackStep; onLike:()=>void; onDislike:()=>void; onOpenForms:()=>void; onSkip:()=>void; }
const FeedbackDialog: React.FC<IFeedbackDialogProps> = function(props) {
  if (props.step === 'idle' || props.step === 'done') { return null; }
  return (
    <div className={styles.feedbackBar}>
      {props.step === 'rating' && (<React.Fragment>
        <div className={styles.feedbackQuestion}>Was this conversation helpful?</div>
        <div className={styles.feedbackActions}>
          <button className={styles.feedbackBtn} onClick={props.onLike} title="Yes, helpful"><ThumbLike20Regular/><span>Yes</span></button>
          <button className={styles.feedbackBtn} onClick={props.onDislike} title="No, not helpful"><ThumbDislike20Regular/><span>No</span></button>
          <button className={styles.feedbackSkip} onClick={props.onSkip}>Skip</button>
        </div>
      </React.Fragment>)}
      {props.step === 'forms' && (<React.Fragment>
        <div className={styles.feedbackQuestion}>Thanks! Want to share more details? <em>(optional)</em></div>
        <div className={styles.feedbackActions}>
          <button className={styles.feedbackBtnPrimary} onClick={props.onOpenForms}>Add feedback</button>
          <button className={styles.feedbackSkip} onClick={props.onSkip}>Skip</button>
        </div>
      </React.Fragment>)}
    </div>
  );
};

/* ── LiveUpdateBanner (unchanged) ───────────────────────────────────────────── */
interface ILiveUpdateBannerProps {
  updates:   ILiveUpdate[];
  dismissed: Record<string, boolean>;
  onDismiss: (id: string) => void;
}
const LiveUpdateBanner: React.FC<ILiveUpdateBannerProps> = function(props) {
  let top: ILiveUpdate | null = null;
  for (let i = 0; i < props.updates.length; i++) {
    if (props.updates[i].priority === 'High' && !props.dismissed[props.updates[i].id]) {
      top = props.updates[i]; break;
    }
  }
  if (!top) { return null; }
  const u = top;
  return (
    <div className={styles.liveUpdateBanner}>
      <Warning20Regular className={styles.bannerIcon} />
      <div className={styles.bannerBody}>
        <div className={styles.bannerType}>{u.updateType} · High priority</div>
        <div className={styles.bannerMsg}>{u.message}</div>
        {u.affectedSession ? <div className={styles.bannerMeta}>Session: {u.affectedSession}</div> : null}
      </div>
      <button className={styles.bannerDismiss} onClick={function() { props.onDismiss(u.id); }} title="Dismiss" aria-label="Dismiss">
        <Dismiss20Regular />
      </button>
    </div>
  );
};

/* ── UpdatesTab (unchanged) ─────────────────────────────────────────────────── */
const UpdatesTab: React.FC<{updates: ILiveUpdate[]}> = function(props) {
  if (props.updates.length === 0) {
    return (<div className={styles.emptyTab}><Alert20Regular className={styles.emptyIcon}/><div className={styles.emptyText}>No active updates at this time.</div></div>);
  }
  return (
    <div className={styles.updatesTab}>
      {props.updates.map(function(u: ILiveUpdate) {
        const dotClass = styles.updateDot + ' ' + (u.priority === 'High' ? styles.updateDotHigh : u.priority === 'Medium' ? styles.updateDotMedium : styles.updateDotLow);
        const showRole = !!u.audienceRole && u.audienceRole.toLowerCase() !== 'all';
        const roleClass = styles.updateRole + ' ' + ((u.audienceRole || '').toLowerCase() === 'speaker' ? styles.updateRoleSpeaker : styles.updateRoleAttendee);
        const expiry = u.expirationTime ? u.expirationTime.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
        return (
          <div key={u.id} className={styles.updateRow}>
            <div className={dotClass} />
            <div className={styles.updateContent}>
              <div className={styles.updateTitle}>{u.title}</div>
              {u.message ? <div className={styles.updateMsg}>{u.message}</div> : null}
              <div className={styles.updateMeta}>
                <span className={styles.updateType}>{u.updateType}</span>
                {showRole ? <span className={roleClass}>{u.audienceRole}</span> : null}
                {expiry ? <span className={styles.updateExpiry}>Until {expiry}</span> : null}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
};

/* ── NearbyTab (unchanged) ──────────────────────────────────────────────────── */
const NearbyTab: React.FC<{recommendations: IRecommendation[]; geocodeContext: string}> = function(props) {
  if (props.recommendations.length === 0) {
    return (<div className={styles.emptyTab}><Map20Regular className={styles.emptyIcon}/><div className={styles.emptyText}>No recommendations configured yet.</div></div>);
  }
  return (
    <div className={styles.nearbyTab}>
      {props.geocodeContext ? (
        <div className={styles.geocodeLabel}><Location20Regular className={styles.geocodeLabelIcon}/><span>Near: {props.geocodeContext}</span></div>
      ) : null}
      {props.recommendations.map(function(r: IRecommendation) {
        const safeLink = sanitizeHttpUrl(r.mapLink);
        return (
          <div key={r.id} className={styles.recRow}>
            <div className={styles.recIconWrap}>{getCategoryIcon(r.category)}</div>
            <div className={styles.recInfo}>
              <div className={styles.recName}>{r.title}</div>
              <div className={styles.recDetail}>
                {r.distanceFromHotel > 0 ? <span>{r.distanceFromHotel.toFixed(1)} mi</span> : null}
                {r.walkingDistance ? <span className={styles.recBadgeWalk}><Tag20Regular className={styles.recBadgeIcon}/>Walk</span> : null}
                {r.eveningFriendly ? <span className={styles.recBadgeEve}><WeatherMoon20Regular className={styles.recBadgeIcon}/>Evening</span> : null}
                {r.category ? <span className={styles.recCategory}>{r.category}</span> : null}
              </div>
              {r.description ? <div className={styles.recDesc}>{r.description}</div> : null}
            </div>
            {safeLink ? (
              <a href={safeLink} target="_blank" rel="noopener noreferrer" className={styles.recMapLink} title={'Open map for ' + r.title} aria-label={'Open map for ' + r.title}>
                <Map20Regular />
              </a>
            ) : null}
          </div>
        );
      })}
    </div>
  );
};

/* ── Helper: copy IChatMessage with new text (unchanged) ─────────────────────── */
function copyMessageWithText(msg: IChatMessage, newText: string): IChatMessage {
  return { id: msg.id, sender: msg.sender, contentType: msg.contentType, text: newText, cardPayload: msg.cardPayload, timestamp: msg.timestamp, suggestedActions: msg.suggestedActions };
}

/* ── helper to strip suggestedActions from all messages ──────────────────────
   Called when the user sends any message so stale chips disappear.           */
function clearAllSuggestedActions(msgs: IChatMessage[]): IChatMessage[] {
  let changed = false;
  const next = msgs.map(function(m) {
    if (m.suggestedActions && m.suggestedActions.length > 0) {
      changed = true;
      return { id: m.id, sender: m.sender, contentType: m.contentType, text: m.text, cardPayload: m.cardPayload, timestamp: m.timestamp };
    }
    return m;
  });
  return changed ? next : msgs;
}

/* ════════════════════════════════════════════════════════════════════════════
   ChatWidget component
   ════════════════════════════════════════════════════════════════════════════ */
export const ChatWidget: React.FC<IChatWidgetProps> = function(props) {
  const spHttpClient      = props.spHttpClient;
  const hubSiteUrl        = props.hubSiteUrl;
  const listTitle         = props.listTitle;
  const defaultTitleMatch = props.defaultTitleMatch;
  const siteId            = props.siteId;
  const siteUrl           = props.siteUrl;
  const userRole          = props.userRole || 'All';

  // ── Original state (all unchanged) ─────────────────────────────────────────
  const [open, setOpen]               = React.useState(false);
  const [closing, setClosing]         = React.useState(false);
  const [maximized, setMaximized]     = React.useState(false);
  const [loading, setLoading]         = React.useState(false);
  const [sending, setSending]         = React.useState(false);
  const [inputText, setInputText]     = React.useState('');
  const [messages, setMessages]       = React.useState<IChatMessage[]>([]);
  const [config, setConfig]           = React.useState<IChatConfig | null>(null);
  const [allowed, setAllowed]         = React.useState<boolean | null>(null);
  const [errorMsg, setErrorMsg]       = React.useState('');
  const [showLabel, setShowLabel]     = React.useState(false);
  const [labelText, setLabelText]     = React.useState('');
  const [prompts, setPrompts]         = React.useState<ISuggestedPrompt[]>([]);
  const [showPrompts, setShowPrompts] = React.useState(true);
  const prevCountRef = React.useRef<number>(0);
  const [refsOpen, setRefsOpen]       = React.useState<Record<string, boolean>>({});
  const [feedbackStep, setFeedbackStep] = React.useState<FeedbackStep>('idle');
  const pendingActionRef = React.useRef<'close'|'reset'|null>(null);
  const proxyRef         = React.useRef<ChatProxyService | null>(null);
  const messagesEndRef   = React.useRef<HTMLDivElement>(null);
  const textareaRef      = React.useRef<HTMLTextAreaElement>(null);
  const messageRefs      = React.useRef<Record<string, HTMLDivElement | null>>({});

  // ── Extended state (unchanged) ──────────────────────────────────────────────
  const [activeTab, setActiveTab]             = React.useState<ActiveTab>('chat');
  const [liveUpdates, setLiveUpdates]         = React.useState<ILiveUpdate[]>([]);
  const [recommendations, setRecommendations] = React.useState<IRecommendation[]>([]);

  // ── dismissed state persists to sessionStorage ────────────────────────────
  const [dismissed, setDismissed] = React.useState<Record<string, boolean>>(function() {
    try {
      const stored = sessionStorage.getItem('ce_dismissed_updates');
      return stored ? JSON.parse(stored) : {};
    } catch { return {}; }
  });

  // ── helper that persists dismissed state to sessionStorage ─────────────────
  function persistDismiss(id: string): void {
    setDismissed(function(prev) {
      const next: Record<string, boolean> = {};
      const keys = Object.keys(prev);
      for (let k = 0; k < keys.length; k++) { next[keys[k]] = prev[keys[k]]; }
      next[id] = true;
      try { sessionStorage.setItem('ce_dismissed_updates', JSON.stringify(next)); } catch { /* ignore */ }
      return next;
    });
  }

  // ── toggleRefs (unchanged) ──────────────────────────────────────────────────
  function toggleRefs(id: string, forceOpen?: boolean): void {
    setRefsOpen(function(s) {
      const next: Record<string, boolean> = {};
      const keys = Object.keys(s);
      for (let k = 0; k < keys.length; k++) { next[keys[k]] = s[keys[k]]; }
      next[id] = forceOpen !== undefined ? forceOpen : !s[id];
      return next;
    });
  }

  // ── Show-once FAB label (unchanged) ─────────────────────────────────────────
  React.useEffect(function() {
    const shown = sessionStorage.getItem(SHOW_ONCE_KEY);
    if (shown) { return; }
    const t1 = window.setTimeout(function() { setShowLabel(true); }, 1800);
    const t2 = window.setTimeout(function() { setShowLabel(false); sessionStorage.setItem(SHOW_ONCE_KEY, '1'); }, 6500);
    return function() { window.clearTimeout(t1); window.clearTimeout(t2); };
  }, []);

  // ── Auto-scroll (unchanged) ──────────────────────────────────────────────────
  React.useEffect(function() {
    if (messages.length <= prevCountRef.current) { prevCountRef.current = messages.length; return; }
    const lastMsg = messages[messages.length - 1];
    prevCountRef.current = messages.length;
    if (lastMsg && lastMsg.sender === 'bot') {
      const el = messageRefs.current[lastMsg.id] || null;
      if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    }
  }, [messages]);

  // ── Auto-expand refs for citations (unchanged) ───────────────────────────────
  React.useEffect(function() {
    setRefsOpen(function(prev) {
      const next: Record<string, boolean> = {};
      const pk = Object.keys(prev);
      for (let k = 0; k < pk.length; k++) { next[pk[k]] = prev[pk[k]]; }
      let changed = false;
      for (let i = 0; i < messages.length; i++) {
        const m = messages[i];
        if (m.sender !== 'bot' || next[m.id] !== undefined) { continue; }
        if (m.contentType !== 'text' || !m.text || typeof m.text !== 'string') { continue; }
        const ex = extractCitationsFromText(m.text);
        if (ex.citations.length > 0) { next[m.id] = true; changed = true; }
      }
      return changed ? next : prev;
    });
  }, [messages]);

  // ── loadExtendedData (unchanged) ────────────────────────────────────────────
  function loadExtendedData(cfg: IChatConfig): void {
    if (!cfg.enableNotifications && !cfg.enableConnectorConsent) { return; }
    if (!cfg.dataSiteUrl) { return; }

    if (cfg.enableNotifications && cfg.liveUpdatesListTitle) {
      LiveUpdatesService.getUpdates(
        spHttpClient, cfg.dataSiteUrl, cfg.liveUpdatesListTitle, userRole
      ).then(function(updates: ILiveUpdate[]) {
        setLiveUpdates(updates);
      }).catch(function() { setLiveUpdates([]); });
    }

    if (cfg.enableConnectorConsent && cfg.recommendationsListTitle) {
      RecommendationsService.getRecommendations(
        spHttpClient, cfg.dataSiteUrl, cfg.recommendationsListTitle
      ).then(function(recs: IRecommendation[]) {
        setRecommendations(recs);
      }).catch(function() { setRecommendations([]); });
    }
  }

  // ── Initialize on open (unchanged) ──────────────────────────────────────────
  React.useEffect(function() {
    if (!open || proxyRef.current) { return; }
    const proxy = new ChatProxyService(spHttpClient, hubSiteUrl, listTitle, siteId, siteUrl, defaultTitleMatch);
    proxyRef.current = proxy;
    setLoading(true);

    proxy.initialize().then(function(result) {
      setAllowed(result.allowed);
      setConfig(result.config);
      setErrorMsg(result.errorMsg);

      if (result.config) {
        setLabelText(result.config.showOnceText || result.config.agentDisplayName);
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const raw: any[] = JSON.parse(result.config.suggestedPromptsJson);
          if (Array.isArray(raw)) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            setPrompts(raw.map(function(item: any) {
              return { title: item.title||'', subtitle: item.subtitle||item.description||'', icon: item.icon||'', prompt: item.prompt||item.title||'' };
            }));
          }
        } catch { setPrompts([]); }

        loadExtendedData(result.config);
      }

      if (result.welcome && result.welcome.length > 0) {
        const sanitized = result.welcome.map(function(msg: IChatMessage) {
          if (msg.sender === 'bot' && msg.contentType === 'text' && typeof msg.text === 'string') {
            return copyMessageWithText(msg, stripUsernameFromWelcome(msg.text));
          }
          return msg;
        });
        setMessages(sanitized);
        setShowPrompts(true);
      }

      setLoading(false);
      window.setTimeout(function() { if (textareaRef.current) { textareaRef.current.focus(); } }, 100);

    }).catch(function(err: unknown) {
      setAllowed(false);
      setErrorMsg('[ERR-UNCAUGHT] ' + ((err instanceof Error) ? err.message : String(err)));
      setLoading(false);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  // ── Feedback helpers (unchanged) ─────────────────────────────────────────────
  function hasFeedbackUrl(): boolean {
    return !!(config && config.msFormsUrl && sanitizeHttpUrl(config.msFormsUrl));
  }
  function hasMeaningfulConvo(): boolean {
    for (let i = 0; i < messages.length; i++) { if (messages[i].sender === 'user') { return true; } }
    return false;
  }
  function executePendingAction(): void {
    const action = pendingActionRef.current;
    pendingActionRef.current = null;
    setFeedbackStep('idle');
    if (action === 'close') {
      setClosing(true);
      window.setTimeout(function() { setOpen(false); setClosing(false); }, 200);
    } else if (action === 'reset') {
      if (!proxyRef.current) { return; }
      setMessages([]); setSending(true);
      proxyRef.current.reset().then(function(w: IChatMessage[]) {
        const sanitized = w.map(function(msg: IChatMessage) {
          if (msg.sender === 'bot' && msg.contentType === 'text' && typeof msg.text === 'string') {
            return copyMessageWithText(msg, stripUsernameFromWelcome(msg.text));
          }
          return msg;
        });
        setMessages(sanitized);
        setShowPrompts(true);
        setActiveTab('chat');
        if (config) { loadExtendedData(config); }
        setSending(false);
        if (textareaRef.current) { textareaRef.current.focus(); }
      }).catch(function() { setSending(false); });
    }
  }
  function requestFeedback(action: 'close'|'reset'): void {
    if (!hasFeedbackUrl() || !hasMeaningfulConvo()) { pendingActionRef.current = action; executePendingAction(); return; }
    pendingActionRef.current = action; setFeedbackStep('rating');
  }
  function handleFeedbackRating(_liked: boolean): void { setFeedbackStep('forms'); }
  function handleOpenForms(): void {
    if (config && config.msFormsUrl) { const safe = sanitizeHttpUrl(config.msFormsUrl); if (safe) { window.open(safe, '_blank', 'noopener,noreferrer'); } }
    executePendingAction();
  }
  function handleFeedbackSkip(): void { executePendingAction(); }
  function handleOpen(): void { setClosing(false); setOpen(true); }
  function handleClose(): void {
    if (feedbackStep !== 'idle') { setFeedbackStep('done'); setClosing(true); window.setTimeout(function() { setOpen(false); setClosing(false); setFeedbackStep('idle'); }, 200); return; }
    requestFeedback('close');
  }
  function handleReset(): void {
    if (feedbackStep !== 'idle') { setFeedbackStep('idle'); pendingActionRef.current = 'reset'; executePendingAction(); return; }
    requestFeedback('reset');
  }

  // ── handleSend (unchanged) ──────────────────────────────────────────────────
  function handleSend(text: string): void {
    const trimmed = text.trim();
    if (!trimmed || sending || !proxyRef.current) { return; }

    setMessages(function(prev) {
      const cleared = clearAllSuggestedActions(prev);
      const userMsg: IChatMessage = {
        id: Math.random().toString(36).slice(2),
        sender: 'user',
        contentType: 'text',
        text: trimmed,
        timestamp: new Date(),
      };
      return cleared.concat([userMsg]);
    });

    setInputText('');
    setShowPrompts(false);
    setActiveTab('chat');
    if (feedbackStep !== 'idle') { setFeedbackStep('idle'); pendingActionRef.current = null; }
    setSending(true);
    if (textareaRef.current) { textareaRef.current.style.height = 'auto'; textareaRef.current.focus(); }

    proxyRef.current.send(trimmed).then(function(replies: IChatMessage[]) {
      setMessages(function(prev) { return prev.concat(replies); });
      setSending(false);
      if (textareaRef.current) { textareaRef.current.focus(); }
    }).catch(function() {
      setSending(false);
      if (textareaRef.current) { textareaRef.current.focus(); }
    });
  }

  function handleKeyDown(e: React.KeyboardEvent<HTMLTextAreaElement>): void {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(inputText); }
  }
  function handleTextareaChange(e: React.ChangeEvent<HTMLTextAreaElement>): void {
    setInputText(e.target.value);
    const ta = e.target; ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight, 100) + 'px';
  }

  function handleDismissBanner(id: string): void { persistDismiss(id); }

  // ── Derived values ───────────────────────────────────────────────────────────
  const displayName = config ? config.agentDisplayName : 'Assistant';
  const agentIconUrl = config ? config.agentIconURL : undefined;
  const panelClass  = [styles.chatPanel, maximized ? styles.maximized : '', closing ? styles.closing : ''].filter(Boolean).join(' ');

  const showUpdatesTab = !!(config && config.enableNotifications && config.dataSiteUrl && config.liveUpdatesListTitle);
  const showNearbyTab  = !!(config && config.enableConnectorConsent && config.dataSiteUrl && config.recommendationsListTitle);
  const showTabs       = showUpdatesTab || showNearbyTab;
  let unreadCount = 0;
  let hasHighPriority = false;
  for (let i = 0; i < liveUpdates.length; i++) {
    if (!dismissed[liveUpdates[i].id]) {
      unreadCount++;
      if (liveUpdates[i].priority === 'High') { hasHighPriority = true; }
    }
  }

  /* ── Render ─────────────────────────────────────────────────────────────────── */
  return (
    <React.Fragment>

      {/* FAB */}
      <button className={styles.fab} onClick={open ? handleClose : handleOpen} aria-label="Toggle chat" title={displayName}>
        <div className={styles.fabInner}>
          <div className={styles.fabPulse} />
          <BotAvatarIcon iconUrl={agentIconUrl} className={styles.fabIcon} />
          {!open && showUpdatesTab && unreadCount > 0 ? (
            <div className={styles.fabNotifBadge} aria-label={unreadCount + ' unread updates'}>
              {unreadCount > 9 ? '9+' : unreadCount}
            </div>
          ) : null}
        </div>
        {!open && labelText ? (
          <div className={styles.fabLabel + (showLabel ? ' ' + styles.visible : '')}>{labelText}</div>
        ) : null}
      </button>

      {/* Panel */}
      {open ? (
        <div className={panelClass} role="dialog" aria-label={displayName + ' chat'}>

          {/* Header */}
          <div className={styles.chatHeader}>
            <div className={styles.headerLeft}>
              <div className={styles.headerAvatar}><BotAvatarIcon iconUrl={agentIconUrl} className={styles.headerAvatarIcon} /></div>
              <div className={styles.headerInfo}>
                <div className={styles.headerTitle}>{displayName}</div>
                <div className={styles.headerSubtitle}>{loading ? 'Connecting…' : (allowed ? 'Powered by Copilot Studio' : '')}</div>
              </div>
            </div>
            <div className={styles.headerActions}>
              <button className={styles.headerBtn} onClick={handleReset} title="New conversation" disabled={loading || !allowed}><ArrowCounterclockwise20Regular /></button>
              <button className={styles.headerBtn} onClick={function() { setMaximized(function(m) { return !m; }); }} title={maximized ? 'Restore' : 'Expand'}>{maximized ? <ArrowMinimize20Regular /> : <ArrowMaximize20Regular />}</button>
              <button className={styles.headerBtnClose} onClick={handleClose} title="Close"><Dismiss20Regular /></button>
            </div>
          </div>

          {/* Tab bar */}
          {showTabs && !loading && allowed ? (
            <div className={styles.tabBar} role="tablist">
              <button role="tab" aria-selected={activeTab === 'chat'} className={styles.tabBtn + (activeTab === 'chat' ? ' ' + styles.tabBtnActive : '')} onClick={function() { setActiveTab('chat'); }}>
                <Chat20Regular className={styles.tabIcon} /><span>Chat</span>
              </button>
              {showUpdatesTab ? (
                <button role="tab" aria-selected={activeTab === 'updates'}
                  className={styles.tabBtn + (activeTab === 'updates' ? ' ' + styles.tabBtnActive : '') + (hasHighPriority ? ' ' + styles.tabBtnAlert : '')}
                  onClick={function() {
                    setActiveTab('updates');
                    for (let j = 0; j < liveUpdates.length; j++) {
                      persistDismiss(liveUpdates[j].id);
                    }
                  }}>
                  <Alert20Regular className={styles.tabIcon} /><span>Updates</span>
                  {unreadCount > 0 ? <span className={styles.tabBadge}>{unreadCount}</span> : null}
                </button>
              ) : null}
              {showNearbyTab ? (
                <button role="tab" aria-selected={activeTab === 'nearby'} className={styles.tabBtn + (activeTab === 'nearby' ? ' ' + styles.tabBtnActive : '')} onClick={function() { setActiveTab('nearby'); }}>
                  <Map20Regular className={styles.tabIcon} /><span>Nearby</span>
                </button>
              ) : null}
            </div>
          ) : null}

          {/* Body */}
          {loading ? (
            <div className={styles.loadingState}>
              <div className={styles.loadingDot} /><div className={styles.loadingDot} /><div className={styles.loadingDot} />
            </div>
          ) : !allowed ? (
            <div className={styles.blockedState}>
              <LockClosed24Regular className={styles.blockedIcon} />
              <div className={styles.blockedTitle}>
                {errorMsg.indexOf('[ERR-CONFIG]')!==-1?'Configuration not found':errorMsg.indexOf('[ERR-DISABLED]')!==-1?'Chatbot disabled':errorMsg.indexOf('[ERR-RBAC]')!==-1?'Access restricted':errorMsg.indexOf('[ERR-MSAL]')!==-1?'Authentication error':errorMsg.indexOf('[ERR-LIST-THROW]')!==-1?'List connection error':errorMsg.indexOf('[ERR-UNCAUGHT]')!==-1?'Unexpected error':'Access restricted'}
              </div>
              <div className={styles.blockedDesc}>{errorMsg}</div>
            </div>
          ) : (
            <React.Fragment>

              {activeTab === 'updates' && showUpdatesTab ? <UpdatesTab updates={liveUpdates} /> : null}

              {activeTab === 'nearby' && showNearbyTab ? (
                <NearbyTab recommendations={recommendations} geocodeContext={config ? (config.geocodeContext || '') : ''} />
              ) : null}

              {activeTab === 'chat' ? (
                <React.Fragment>

                  {showUpdatesTab ? <LiveUpdateBanner updates={liveUpdates} dismissed={dismissed} onDismiss={handleDismissBanner} /> : null}

                  {/* Suggested prompts */}
                  {showPrompts && messages.length <= 2 && prompts.length > 0 ? (
                    <div className={styles.promptsSection}>
                      <div className={styles.promptsLabel}><Sparkle20Regular className={styles.promptsLabelIcon} /><span>Suggested prompts</span></div>
                      <div className={styles.suggestedPrompts}>
                        {prompts.slice(0, 6).map(function(p: ISuggestedPrompt, i: number) {
                          return (
                            <button key={i} className={styles.promptChip} onClick={function() { handleSend(p.prompt); }}>
                              <div className={styles.promptIconWrap}>{getPromptIcon(p.icon || p.title)}</div>
                              <div className={styles.promptText}><div className={styles.promptTitle}>{p.title}</div></div>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  ) : null}

                  {/* Messages */}
                  <div className={styles.messagesContainer}>
                    {messages.map(function(msg: IChatMessage) {
                      const isUser = msg.sender === 'user';
                      const rowClass = styles.messageRow + (isUser ? ' ' + styles.userRow : '');
                      const bubClass = styles.bubble + ' ' + (isUser ? styles.userBubble : msg.sender === 'system' ? styles.errorBubble : styles.botBubble);
                      return (
                        <div key={msg.id} ref={function(el) { messageRefs.current[msg.id] = el; }} className={rowClass}>
                          {msg.sender === 'bot' ? <div className={styles.msgAvatar}><BotAvatarIcon iconUrl={agentIconUrl} className={styles.msgAvatarIcon} /></div> : null}
                          <div className={styles.msgContent}>
                            {msg.contentType === 'adaptiveCard' && msg.cardPayload ? (
                              <div className={styles.bubble + ' ' + styles.botBubble}><AdaptiveCardHost payload={msg.cardPayload} /></div>
                            ) : (
                              <div className={bubClass}>
                                {(function() {
                                  const ex = extractCitationsFromText(msg.text || '');
                                  const byIdx: Record<number, ICitation> = {};
                                  for (let ci = 0; ci < ex.citations.length; ci++) { byIdx[ex.citations[ci].index] = ex.citations[ci]; }
                                  const onCC = function(n: number) { toggleRefs(msg.id, true); const el = messageRefs.current[msg.id] || null; if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); } };
                                  return (
                                    <React.Fragment>
                                      {msg.sender === 'bot' && msg.text ? renderMarkdown(ex.cleanText || msg.text, { citationsByIndex: byIdx, onCitationClick: onCC }) : msg.text}
                                      {ex.citations.length > 0 ? (
                                        <div className={styles.refsSection}>
                                          {(function() {
                                            const isOpen = refsOpen[msg.id] !== undefined ? refsOpen[msg.id] : true;
                                            return (
                                              <React.Fragment>
                                                <button type="button" className={styles.refsToggle} onClick={function() { toggleRefs(msg.id); }} aria-expanded={isOpen}>
                                                  {ex.citations.length} reference{ex.citations.length > 1 ? 's' : ''}{' '}<span className={styles.chev}>{refsOpen[msg.id] ? '▾' : '▸'}</span>
                                                </button>
                                                {refsOpen[msg.id] ? (
                                                  <div className={styles.refsList}>
                                                    {ex.citations.slice().sort(function(a, b) { return a.index - b.index; }).map(function(c: ICitation) {
                                                      return (
                                                        <a key={c.index} href={sanitizeHttpUrl(c.url) || '#'} target="_blank" rel="noopener noreferrer" className={styles.refChip}>
                                                          <span className={styles.refIndex}>{c.index}</span>
                                                          <span className={styles.refLabel}>{niceLabelFromUrl(c.url, c.title)}</span>
                                                        </a>
                                                      );
                                                    })}
                                                  </div>
                                                ) : null}
                                              </React.Fragment>
                                            );
                                          })()}
                                        </div>
                                      ) : null}
                                    </React.Fragment>
                                  );
                                })()}
                              </div>
                            )}

                            {/* suggested action chips below this bot bubble */}
                            {msg.sender === 'bot' && msg.suggestedActions && msg.suggestedActions.length > 0 ? (
                              <div className={styles.suggestedActionsRow}>
                                {msg.suggestedActions.map(function(action: ISuggestedAction, ai: number) {
                                  return (
                                    <button
                                      key={ai}
                                      className={styles.suggestedActionChip}
                                      onClick={function() { handleSend(action.value); }}
                                      disabled={sending}
                                      title={action.title}
                                    >
                                      {action.title}
                                    </button>
                                  );
                                })}
                              </div>
                            ) : null}

                            <div className={styles.msgTime + (isUser ? ' ' + styles.msgTimeUser : '')}>{formatTime(msg.timestamp)}</div>
                          </div>
                        </div>
                      );
                    })}
                    {sending ? (
                      <div className={styles.typingRow}>
                        <div className={styles.msgAvatar}><BotAvatarIcon iconUrl={agentIconUrl} className={styles.msgAvatarIcon} /></div>
                        <div className={styles.typingBubble}><span /><span /><span /></div>
                      </div>
                    ) : null}
                    <div ref={messagesEndRef} />
                  </div>

                  <FeedbackDialog step={feedbackStep} onLike={function() { handleFeedbackRating(true); }} onDislike={function() { handleFeedbackRating(false); }} onOpenForms={handleOpenForms} onSkip={handleFeedbackSkip} />

                  <div className={styles.inputArea}>
                    <div className={styles.inputWrap}>
                      <textarea ref={textareaRef} className={styles.textInput} placeholder="Ask a question…" value={inputText} onChange={handleTextareaChange} onKeyDown={handleKeyDown} rows={1} disabled={sending} />
                      <button className={styles.sendBtn} onClick={function() { handleSend(inputText); }} disabled={!inputText.trim() || sending} title="Send"><Send20Filled /></button>
                    </div>
                    <div className={styles.inputFooter}>Powered by Microsoft Copilot Studio</div>
                  </div>

                </React.Fragment>
              ) : null}

            </React.Fragment>
          )}

        </div>
      ) : null}

    </React.Fragment>
  );
};