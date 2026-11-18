import * as React from 'react';
import { BotSparkle24Filled, ArrowCounterclockwise20Regular, ArrowMaximize20Regular, ArrowMinimize20Regular, Dismiss20Regular, Send20Filled, LockClosed24Regular, Settings20Regular, Person20Regular, ShieldLock20Regular, ClipboardTaskListLtr20Regular, Wrench20Regular, DocumentBulletList20Regular, ChatMultiple20Regular, Sparkle20Regular, ThumbLike20Regular, ThumbDislike20Regular, Chat20Regular, Alert20Regular, Map20Regular, Warning20Regular, Location20Regular, Tag20Regular, WeatherMoon20Regular, MusicNote120Regular, Food20Regular, Building20Regular, } from '@fluentui/react-icons';
import { ChatProxyService } from '../services/ChatProxyservices';
import { LiveUpdatesService } from '../services/Liveupdatesservice';
import { RecommendationsService } from '../services/Recommendationsservice';
import { AdaptiveCardHost } from './AdaptiveCardHost';
import styles from './ChatWidget.module.scss';
// ─── Constants ────────────────────────────────────────────────────────────────
var SHOW_ONCE_KEY = 'ce_label_shown';
function formatTime(d) {
    return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}
function stripUsernameFromWelcome(text) {
    if (!text)
        return text;
    return text.replace(/^(Hello|Hi|Hey|Greetings|Welcome)(?:,?\s+)([A-Z][a-z]+)([,!.]?)/, '$1$3');
}
function sanitizeHttpUrl(raw) {
    if (!raw)
        return null;
    var t = raw.trim();
    return /^https?:\/\//i.test(t) ? t : null;
}
function niceLabelFromUrl(u, explicitTitle) {
    if (explicitTitle)
        return explicitTitle;
    try {
        var url = new URL(u);
        var parts = url.pathname.split('/').filter(function (s) { return s.length > 0; });
        var last = parts.length > 0 ? parts[parts.length - 1] : url.hostname;
        var base = decodeURIComponent(last).replace(/\.[a-z0-9]+$/i, '');
        var label = base || url.hostname;
        return label.length > 36 ? label.slice(0, 33) + '…' : label;
    }
    catch (_a) {
        return u.length > 36 ? u.slice(0, 33) + '…' : u;
    }
}
function extractCitationsFromText(rawText) {
    var lines = (rawText || '').split('\n');
    var body = [];
    var citObj = {};
    var rBracketSame = /^\s*\[(\d+)\]\s*:\s*(https?:\S+)(?:\s+"([^"]+)")?\s*$/i;
    var rBracketAlone = /^\s*\[(\d+)\]\s*:\s*$/i;
    var rNumberDot = /^\s*(\d+)[.)]\s+(https?:\S+)(?:\s+"([^"]+)")?\s*$/i;
    var rUrlOnly = /^\s*(https?:\S+)(?:\s+"([^"]+)")?\s*$/i;
    var rRefsHeading = /^\s*#{0,3}\s*references?:?\s*$/i;
    var i = 0;
    var inRefsBlock = false;
    var maxIdx = 0;
    while (i < lines.length) {
        var line = lines[i];
        if (rRefsHeading.test(line)) {
            inRefsBlock = true;
            i++;
            continue;
        }
        var mSame = rBracketSame.exec(line);
        if (mSame) {
            var idx = parseInt(mSame[1], 10);
            var url = sanitizeHttpUrl(mSame[2]);
            if (url) {
                citObj[idx] = { index: idx, url: url, title: mSame[3] };
                if (idx > maxIdx)
                    maxIdx = idx;
            }
            inRefsBlock = true;
            i++;
            continue;
        }
        var mAlone = rBracketAlone.exec(line);
        if (mAlone) {
            var idx = parseInt(mAlone[1], 10);
            var j = i + 1;
            while (j < lines.length && lines[j].trim() === '')
                j++;
            if (j < lines.length) {
                var mNext = rUrlOnly.exec(lines[j]);
                if (mNext) {
                    var url = sanitizeHttpUrl(mNext[1]);
                    if (url) {
                        citObj[idx] = { index: idx, url: url, title: mNext[2] };
                        if (idx > maxIdx)
                            maxIdx = idx;
                    }
                    i = j + 1;
                    inRefsBlock = true;
                    continue;
                }
            }
            inRefsBlock = true;
            i++;
            continue;
        }
        var mNum = rNumberDot.exec(line);
        if (mNum) {
            var idx = parseInt(mNum[1], 10);
            var url = sanitizeHttpUrl(mNum[2]);
            if (url) {
                citObj[idx] = { index: idx, url: url, title: mNum[3] };
                if (idx > maxIdx)
                    maxIdx = idx;
            }
            inRefsBlock = true;
            i++;
            continue;
        }
        if (inRefsBlock) {
            var mUrl = rUrlOnly.exec(line);
            if (mUrl) {
                var ni = maxIdx + 1;
                var url = sanitizeHttpUrl(mUrl[1]);
                if (url) {
                    citObj[ni] = { index: ni, url: url, title: mUrl[2] };
                    maxIdx = ni;
                }
                i++;
                continue;
            }
        }
        body.push(line);
        i++;
    }
    var citations = [];
    var keys = Object.keys(citObj);
    for (var k = 0; k < keys.length; k++) {
        citations.push(citObj[parseInt(keys[k], 10)]);
    }
    citations.sort(function (a, b) { return a.index - b.index; });
    return { cleanText: body.join('\n').trim(), citations: citations };
}
function renderMarkdown(text, opts) {
    if (!text) {
        return null;
    }
    var lines = text.split('\n');
    var nodes = [];
    var listItems = [];
    var listType = null;
    var paraLines = [];
    var key = 0;
    function flushList() { if (!listItems.length)
        return; if (listType === 'ol') {
        nodes.push(React.createElement("ol", { key: 'ol-' + key++, className: styles.mdOl }, listItems));
    }
    else {
        nodes.push(React.createElement("ul", { key: 'ul-' + key++, className: styles.mdUl }, listItems));
    } listItems = []; listType = null; }
    function flushPara() { if (!paraLines.length)
        return; nodes.push(React.createElement("p", { key: 'p-' + key++, className: styles.mdP }, inlineFormat(paraLines.join(' ')))); paraLines = []; }
    function inlineFormat(raw) {
        var parts = [];
        var remaining = raw;
        var idx = 0;
        var _loop_1 = function () {
            var boldM = remaining.match(/^(.*?)\*\*(.+?)\*\*(.*)/);
            var italM = remaining.match(/^(.*?)\*(.+?)\*(.*)/);
            var codeM = remaining.match(/^(.*?)`(.+?)`(.*)/);
            var linkM = remaining.match(/^(.*?)\[(.+?)\]\((.+?)\)(.*)/);
            var citeM = remaining.match(/^(.*?)\[(\d+)\](.*)/);
            // ── NEW: bare URL and autolink patterns ────────────────────────────────
            // autolink:  <https://example.com>     → linkify, render without brackets
            // bare url:  https://example.com       → linkify in place
            // Stops on whitespace and chars that almost always end a URL in prose.
            // FIX: added \] and \[ to the excluded class below so a bare URL stops
            // cleanly before a following [1] citation marker instead of consuming
            // the bracket and digit into the href/link text (e.g. "...page[1] [2]."
            // was previously linking "...page[1" and breaking the first citation).
            var autolinkM = remaining.match(/^(.*?)<(https?:\/\/[^\s>]+)>(.*)/);
            var urlM = remaining.match(/^(.*?)(https?:\/\/[^\s<>"')\]\[]+)(.*)/);
            var matches = [];
            if (boldM && boldM[1] !== undefined)
                matches.push({ pos: boldM[1].length, type: 'bold', m: boldM });
            if (italM && italM[1] !== undefined)
                matches.push({ pos: italM[1].length, type: 'italic', m: italM });
            if (codeM && codeM[1] !== undefined)
                matches.push({ pos: codeM[1].length, type: 'code', m: codeM });
            if (linkM && linkM[1] !== undefined)
                matches.push({ pos: linkM[1].length, type: 'link', m: linkM });
            if (citeM && citeM[1] !== undefined)
                matches.push({ pos: citeM[1].length, type: 'cite', m: citeM });
            if (autolinkM && autolinkM[1] !== undefined)
                matches.push({ pos: autolinkM[1].length, type: 'autolink', m: autolinkM });
            if (urlM && urlM[1] !== undefined)
                matches.push({ pos: urlM[1].length, type: 'url', m: urlM });
            if (!matches.length) {
                parts.push(remaining);
                return "break";
            }
            matches.sort(function (a, b) { return a.pos - b.pos; });
            var best = matches[0];
            if (best.m[1])
                parts.push(best.m[1]);
            if (best.type === 'bold') {
                parts.push(React.createElement("strong", { key: idx++ }, best.m[2]));
                remaining = best.m[3];
            }
            else if (best.type === 'italic') {
                parts.push(React.createElement("em", { key: idx++ }, best.m[2]));
                remaining = best.m[3];
            }
            else if (best.type === 'code') {
                parts.push(React.createElement("code", { key: idx++, className: styles.mdCode }, best.m[2]));
                remaining = best.m[3];
            }
            else if (best.type === 'link') {
                var raw2 = best.m[3];
                var safe = (raw2.indexOf('http://') === 0 || raw2.indexOf('https://') === 0) ? raw2 : '#';
                parts.push(React.createElement("a", { key: idx++, href: safe, target: "_blank", rel: "noopener noreferrer", className: styles.mdLink }, best.m[2]));
                remaining = best.m[4] || '';
            }
            else if (best.type === 'cite') {
                var n_1 = parseInt(best.m[2], 10);
                var cit = (opts && opts.citationsByIndex) ? opts.citationsByIndex[n_1] : undefined;
                if (cit && cit.url) {
                    var safe = sanitizeHttpUrl(cit.url) || '#';
                    var ct = cit.title ? 'Open ' + cit.title : 'Open reference ' + String(n_1);
                    var oc = function () { if (opts && opts.onCitationClick)
                        opts.onCitationClick(n_1); };
                    parts.push(React.createElement("a", { key: idx++, href: safe, target: "_blank", rel: "noopener noreferrer", className: styles.citeSup, onClick: oc, title: ct }, n_1));
                }
                else {
                    parts.push('[' + best.m[2] + ']');
                }
                remaining = best.m[3] || '';
            }
            // ── NEW: autolink branch ────────────────────────────────────────────────
            else if (best.type === 'autolink') {
                var url = best.m[2];
                var safe = sanitizeHttpUrl(url) || '#';
                parts.push(React.createElement("a", { key: idx++, href: safe, target: "_blank", rel: "noopener noreferrer", className: styles.mdLink }, url));
                remaining = best.m[3] || '';
            }
            // ── NEW: bare URL branch ────────────────────────────────────────────────
            else if (best.type === 'url') {
                var url = best.m[2];
                var tail = '';
                // Strip trailing punctuation that almost never belongs to a URL,
                // e.g. "see https://example.com." or "(https://x.com)"
                while (url.length > 0 && /[.,;:!?)\]}'"`]/.test(url.charAt(url.length - 1))) {
                    tail = url.charAt(url.length - 1) + tail;
                    url = url.slice(0, -1);
                }
                var safe = sanitizeHttpUrl(url) || '#';
                parts.push(React.createElement("a", { key: idx++, href: safe, target: "_blank", rel: "noopener noreferrer", className: styles.mdLink }, url));
                remaining = tail + (best.m[3] || '');
            }
        };
        while (remaining.length > 0) {
            var state_1 = _loop_1();
            if (state_1 === "break")
                break;
        }
        return parts.length === 1 ? parts[0] : React.createElement(React.Fragment, null, parts);
    }
    for (var i = 0; i < lines.length; i++) {
        var t = lines[i].trim();
        if (t === '') {
            flushList();
            flushPara();
            continue;
        }
        var hM = t.match(/^(#{1,3})\s+(.+)/);
        if (hM) {
            flushList();
            flushPara();
            var lvl = hM[1].length;
            var hc = inlineFormat(hM[2]);
            if (lvl === 1)
                nodes.push(React.createElement("div", { key: 'h-' + key++, className: styles.mdH1 }, hc));
            else if (lvl === 2)
                nodes.push(React.createElement("div", { key: 'h-' + key++, className: styles.mdH2 }, hc));
            else
                nodes.push(React.createElement("div", { key: 'h-' + key++, className: styles.mdH3 }, hc));
            continue;
        }
        var ulM = t.match(/^[-*]\s+(.+)/);
        if (ulM) {
            flushPara();
            if (listType === 'ol')
                flushList();
            listType = 'ul';
            listItems.push(React.createElement("li", { key: 'li-' + key++, className: styles.mdLi }, inlineFormat(ulM[1])));
            continue;
        }
        var olM = t.match(/^\d+\.\s+(.+)/);
        if (olM) {
            flushPara();
            if (listType === 'ul')
                flushList();
            listType = 'ol';
            listItems.push(React.createElement("li", { key: 'li-' + key++, className: styles.mdLi }, inlineFormat(olM[1])));
            continue;
        }
        flushList();
        paraLines.push(t);
    }
    flushList();
    flushPara();
    return React.createElement(React.Fragment, null, nodes);
}
/* ── Prompt icon map (unchanged) ────────────────────────────────────────────── */
function getPromptIcon(iconName) {
    var n = (iconName || '').toLowerCase();
    if (n.indexOf('process') !== -1 || n.indexOf('gear') !== -1 || n.indexOf('setting') !== -1 || n.indexOf('workflow') !== -1)
        return React.createElement(Settings20Regular, null);
    if (n.indexOf('contact') !== -1 || n.indexOf('people') !== -1 || n.indexOf('poi') !== -1 || n.indexOf('local') !== -1)
        return React.createElement(Person20Regular, null);
    if (n.indexOf('access') !== -1 || n.indexOf('permission') !== -1 || n.indexOf('security') !== -1 || n.indexOf('lock') !== -1 || n.indexOf('password') !== -1)
        return React.createElement(ShieldLock20Regular, null);
    if (n.indexOf('onboard') !== -1 || n.indexOf('checklist') !== -1 || n.indexOf('task') !== -1 || n.indexOf('setup') !== -1)
        return React.createElement(ClipboardTaskListLtr20Regular, null);
    if (n.indexOf('troubl') !== -1 || n.indexOf('wrench') !== -1 || n.indexOf('tool') !== -1 || n.indexOf('fix') !== -1 || n.indexOf('help') !== -1)
        return React.createElement(Wrench20Regular, null);
    if (n.indexOf('polic') !== -1 || n.indexOf('document') !== -1 || n.indexOf('rule') !== -1 || n.indexOf('guidanc') !== -1)
        return React.createElement(DocumentBulletList20Regular, null);
    return React.createElement(ChatMultiple20Regular, null);
}
/* ── Category icon for recommendations (unchanged) ──────────────────────────── */
function getCategoryIcon(category) {
    var c = (category || '').toLowerCase();
    if (c.indexOf('jazz') !== -1 || c.indexOf('music') !== -1) {
        return React.createElement(MusicNote120Regular, null);
    }
    if (c.indexOf('food') !== -1 || c.indexOf('cafe') !== -1 || c.indexOf('restaurant') !== -1) {
        return React.createElement(Food20Regular, null);
    }
    if (c.indexOf('sight') !== -1 || c.indexOf('tour') !== -1 || c.indexOf('walk') !== -1) {
        return React.createElement(Building20Regular, null);
    }
    return React.createElement(Location20Regular, null);
}
var BotAvatarIcon = function (props) {
    var _a = React.useState(false), failed = _a[0], setFailed = _a[1];
    var safeUrl = sanitizeHttpUrl(props.iconUrl);
    if (safeUrl && !failed) {
        return (React.createElement("img", { src: safeUrl, alt: "", className: props.className, onError: function () { setFailed(true); } }));
    }
    return React.createElement(BotSparkle24Filled, { className: props.className });
};
var FeedbackDialog = function (props) {
    if (props.step === 'idle' || props.step === 'done') {
        return null;
    }
    return (React.createElement("div", { className: styles.feedbackBar },
        props.step === 'rating' && (React.createElement(React.Fragment, null,
            React.createElement("div", { className: styles.feedbackQuestion }, "Was this conversation helpful?"),
            React.createElement("div", { className: styles.feedbackActions },
                React.createElement("button", { className: styles.feedbackBtn, onClick: props.onLike, title: "Yes, helpful" },
                    React.createElement(ThumbLike20Regular, null),
                    React.createElement("span", null, "Yes")),
                React.createElement("button", { className: styles.feedbackBtn, onClick: props.onDislike, title: "No, not helpful" },
                    React.createElement(ThumbDislike20Regular, null),
                    React.createElement("span", null, "No")),
                React.createElement("button", { className: styles.feedbackSkip, onClick: props.onSkip }, "Skip")))),
        props.step === 'forms' && (React.createElement(React.Fragment, null,
            React.createElement("div", { className: styles.feedbackQuestion },
                "Thanks! Want to share more details? ",
                React.createElement("em", null, "(optional)")),
            React.createElement("div", { className: styles.feedbackActions },
                React.createElement("button", { className: styles.feedbackBtnPrimary, onClick: props.onOpenForms }, "Add feedback"),
                React.createElement("button", { className: styles.feedbackSkip, onClick: props.onSkip }, "Skip"))))));
};
var LiveUpdateBanner = function (props) {
    var top = null;
    for (var i = 0; i < props.updates.length; i++) {
        if (props.updates[i].priority === 'High' && !props.dismissed[props.updates[i].id]) {
            top = props.updates[i];
            break;
        }
    }
    if (!top) {
        return null;
    }
    var u = top;
    return (React.createElement("div", { className: styles.liveUpdateBanner },
        React.createElement(Warning20Regular, { className: styles.bannerIcon }),
        React.createElement("div", { className: styles.bannerBody },
            React.createElement("div", { className: styles.bannerType },
                u.updateType,
                " \u00B7 High priority"),
            React.createElement("div", { className: styles.bannerMsg }, u.message),
            u.affectedSession ? React.createElement("div", { className: styles.bannerMeta },
                "Session: ",
                u.affectedSession) : null),
        React.createElement("button", { className: styles.bannerDismiss, onClick: function () { props.onDismiss(u.id); }, title: "Dismiss", "aria-label": "Dismiss" },
            React.createElement(Dismiss20Regular, null))));
};
/* ── UpdatesTab (unchanged) ─────────────────────────────────────────────────── */
var UpdatesTab = function (props) {
    if (props.updates.length === 0) {
        return (React.createElement("div", { className: styles.emptyTab },
            React.createElement(Alert20Regular, { className: styles.emptyIcon }),
            React.createElement("div", { className: styles.emptyText }, "No active updates at this time.")));
    }
    return (React.createElement("div", { className: styles.updatesTab }, props.updates.map(function (u) {
        var dotClass = styles.updateDot + ' ' + (u.priority === 'High' ? styles.updateDotHigh : u.priority === 'Medium' ? styles.updateDotMedium : styles.updateDotLow);
        var showRole = !!u.audienceRole && u.audienceRole.toLowerCase() !== 'all';
        var roleClass = styles.updateRole + ' ' + ((u.audienceRole || '').toLowerCase() === 'speaker' ? styles.updateRoleSpeaker : styles.updateRoleAttendee);
        var expiry = u.expirationTime ? u.expirationTime.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
        return (React.createElement("div", { key: u.id, className: styles.updateRow },
            React.createElement("div", { className: dotClass }),
            React.createElement("div", { className: styles.updateContent },
                React.createElement("div", { className: styles.updateTitle }, u.title),
                u.message ? React.createElement("div", { className: styles.updateMsg }, u.message) : null,
                React.createElement("div", { className: styles.updateMeta },
                    React.createElement("span", { className: styles.updateType }, u.updateType),
                    showRole ? React.createElement("span", { className: roleClass }, u.audienceRole) : null,
                    expiry ? React.createElement("span", { className: styles.updateExpiry },
                        "Until ",
                        expiry) : null))));
    })));
};
/* ── NearbyTab (unchanged) ──────────────────────────────────────────────────── */
var NearbyTab = function (props) {
    if (props.recommendations.length === 0) {
        return (React.createElement("div", { className: styles.emptyTab },
            React.createElement(Map20Regular, { className: styles.emptyIcon }),
            React.createElement("div", { className: styles.emptyText }, "No recommendations configured yet.")));
    }
    return (React.createElement("div", { className: styles.nearbyTab },
        props.geocodeContext ? (React.createElement("div", { className: styles.geocodeLabel },
            React.createElement(Location20Regular, { className: styles.geocodeLabelIcon }),
            React.createElement("span", null,
                "Near: ",
                props.geocodeContext))) : null,
        props.recommendations.map(function (r) {
            var safeLink = sanitizeHttpUrl(r.mapLink);
            return (React.createElement("div", { key: r.id, className: styles.recRow },
                React.createElement("div", { className: styles.recIconWrap }, getCategoryIcon(r.category)),
                React.createElement("div", { className: styles.recInfo },
                    React.createElement("div", { className: styles.recName }, r.title),
                    React.createElement("div", { className: styles.recDetail },
                        r.distanceFromHotel > 0 ? React.createElement("span", null,
                            r.distanceFromHotel.toFixed(1),
                            " mi") : null,
                        r.walkingDistance ? React.createElement("span", { className: styles.recBadgeWalk },
                            React.createElement(Tag20Regular, { className: styles.recBadgeIcon }),
                            "Walk") : null,
                        r.eveningFriendly ? React.createElement("span", { className: styles.recBadgeEve },
                            React.createElement(WeatherMoon20Regular, { className: styles.recBadgeIcon }),
                            "Evening") : null,
                        r.category ? React.createElement("span", { className: styles.recCategory }, r.category) : null),
                    r.description ? React.createElement("div", { className: styles.recDesc }, r.description) : null),
                safeLink ? (React.createElement("a", { href: safeLink, target: "_blank", rel: "noopener noreferrer", className: styles.recMapLink, title: 'Open map for ' + r.title, "aria-label": 'Open map for ' + r.title },
                    React.createElement(Map20Regular, null))) : null));
        })));
};
/* ── Helper: copy IChatMessage with new text (unchanged) ─────────────────────── */
function copyMessageWithText(msg, newText) {
    return { id: msg.id, sender: msg.sender, contentType: msg.contentType, text: newText, cardPayload: msg.cardPayload, timestamp: msg.timestamp, suggestedActions: msg.suggestedActions };
}
/* ── helper to strip suggestedActions from all messages ──────────────────────
   Called when the user sends any message so stale chips disappear.           */
function clearAllSuggestedActions(msgs) {
    var changed = false;
    var next = msgs.map(function (m) {
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
export var ChatWidget = function (props) {
    var spHttpClient = props.spHttpClient;
    var hubSiteUrl = props.hubSiteUrl;
    var listTitle = props.listTitle;
    var defaultTitleMatch = props.defaultTitleMatch;
    var siteId = props.siteId;
    var siteUrl = props.siteUrl;
    var userRole = props.userRole || 'All';
    // ── Original state (all unchanged) ─────────────────────────────────────────
    var _a = React.useState(false), open = _a[0], setOpen = _a[1];
    var _b = React.useState(false), closing = _b[0], setClosing = _b[1];
    var _c = React.useState(false), maximized = _c[0], setMaximized = _c[1];
    var _d = React.useState(false), loading = _d[0], setLoading = _d[1];
    var _e = React.useState(false), sending = _e[0], setSending = _e[1];
    var _f = React.useState(''), inputText = _f[0], setInputText = _f[1];
    var _g = React.useState([]), messages = _g[0], setMessages = _g[1];
    var _h = React.useState(null), config = _h[0], setConfig = _h[1];
    var _j = React.useState(null), allowed = _j[0], setAllowed = _j[1];
    var _k = React.useState(''), errorMsg = _k[0], setErrorMsg = _k[1];
    var _l = React.useState(false), showLabel = _l[0], setShowLabel = _l[1];
    var _m = React.useState(''), labelText = _m[0], setLabelText = _m[1];
    var _o = React.useState([]), prompts = _o[0], setPrompts = _o[1];
    var _p = React.useState(true), showPrompts = _p[0], setShowPrompts = _p[1];
    var prevCountRef = React.useRef(0);
    var _q = React.useState({}), refsOpen = _q[0], setRefsOpen = _q[1];
    var _r = React.useState('idle'), feedbackStep = _r[0], setFeedbackStep = _r[1];
    var pendingActionRef = React.useRef(null);
    var proxyRef = React.useRef(null);
    var messagesEndRef = React.useRef(null);
    var textareaRef = React.useRef(null);
    var messageRefs = React.useRef({});
    // ── Extended state (unchanged) ──────────────────────────────────────────────
    var _s = React.useState('chat'), activeTab = _s[0], setActiveTab = _s[1];
    var _t = React.useState([]), liveUpdates = _t[0], setLiveUpdates = _t[1];
    var _u = React.useState([]), recommendations = _u[0], setRecommendations = _u[1];
    // ── dismissed state persists to sessionStorage ────────────────────────────
    var _v = React.useState(function () {
        try {
            var stored = sessionStorage.getItem('ce_dismissed_updates');
            return stored ? JSON.parse(stored) : {};
        }
        catch (_a) {
            return {};
        }
    }), dismissed = _v[0], setDismissed = _v[1];
    // ── helper that persists dismissed state to sessionStorage ─────────────────
    function persistDismiss(id) {
        setDismissed(function (prev) {
            var next = {};
            var keys = Object.keys(prev);
            for (var k = 0; k < keys.length; k++) {
                next[keys[k]] = prev[keys[k]];
            }
            next[id] = true;
            try {
                sessionStorage.setItem('ce_dismissed_updates', JSON.stringify(next));
            }
            catch ( /* ignore */_a) { /* ignore */ }
            return next;
        });
    }
    // ── toggleRefs (unchanged) ──────────────────────────────────────────────────
    function toggleRefs(id, forceOpen) {
        setRefsOpen(function (s) {
            var next = {};
            var keys = Object.keys(s);
            for (var k = 0; k < keys.length; k++) {
                next[keys[k]] = s[keys[k]];
            }
            next[id] = forceOpen !== undefined ? forceOpen : !s[id];
            return next;
        });
    }
    // ── Show-once FAB label (unchanged) ─────────────────────────────────────────
    React.useEffect(function () {
        var shown = sessionStorage.getItem(SHOW_ONCE_KEY);
        if (shown) {
            return;
        }
        var t1 = window.setTimeout(function () { setShowLabel(true); }, 1800);
        var t2 = window.setTimeout(function () { setShowLabel(false); sessionStorage.setItem(SHOW_ONCE_KEY, '1'); }, 6500);
        return function () { window.clearTimeout(t1); window.clearTimeout(t2); };
    }, []);
    // ── Auto-scroll (unchanged) ──────────────────────────────────────────────────
    React.useEffect(function () {
        if (messages.length <= prevCountRef.current) {
            prevCountRef.current = messages.length;
            return;
        }
        var lastMsg = messages[messages.length - 1];
        prevCountRef.current = messages.length;
        if (lastMsg && lastMsg.sender === 'bot') {
            var el = messageRefs.current[lastMsg.id] || null;
            if (el) {
                el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    }, [messages]);
    // ── Auto-expand refs for citations (unchanged) ───────────────────────────────
    React.useEffect(function () {
        setRefsOpen(function (prev) {
            var next = {};
            var pk = Object.keys(prev);
            for (var k = 0; k < pk.length; k++) {
                next[pk[k]] = prev[pk[k]];
            }
            var changed = false;
            for (var i = 0; i < messages.length; i++) {
                var m = messages[i];
                if (m.sender !== 'bot' || next[m.id] !== undefined) {
                    continue;
                }
                if (m.contentType !== 'text' || !m.text || typeof m.text !== 'string') {
                    continue;
                }
                var ex = extractCitationsFromText(m.text);
                if (ex.citations.length > 0) {
                    next[m.id] = true;
                    changed = true;
                }
            }
            return changed ? next : prev;
        });
    }, [messages]);
    // ── loadExtendedData (unchanged) ────────────────────────────────────────────
    function loadExtendedData(cfg) {
        if (!cfg.enableNotifications && !cfg.enableConnectorConsent) {
            return;
        }
        if (!cfg.dataSiteUrl) {
            return;
        }
        if (cfg.enableNotifications && cfg.liveUpdatesListTitle) {
            LiveUpdatesService.getUpdates(spHttpClient, cfg.dataSiteUrl, cfg.liveUpdatesListTitle, userRole).then(function (updates) {
                setLiveUpdates(updates);
            }).catch(function () { setLiveUpdates([]); });
        }
        if (cfg.enableConnectorConsent && cfg.recommendationsListTitle) {
            RecommendationsService.getRecommendations(spHttpClient, cfg.dataSiteUrl, cfg.recommendationsListTitle).then(function (recs) {
                setRecommendations(recs);
            }).catch(function () { setRecommendations([]); });
        }
    }
    // ── Initialize on open (unchanged) ──────────────────────────────────────────
    React.useEffect(function () {
        if (!open || proxyRef.current) {
            return;
        }
        var proxy = new ChatProxyService(spHttpClient, hubSiteUrl, listTitle, siteId, siteUrl, defaultTitleMatch);
        proxyRef.current = proxy;
        setLoading(true);
        proxy.initialize().then(function (result) {
            setAllowed(result.allowed);
            setConfig(result.config);
            setErrorMsg(result.errorMsg);
            if (result.config) {
                setLabelText(result.config.showOnceText || result.config.agentDisplayName);
                try {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    var raw = JSON.parse(result.config.suggestedPromptsJson);
                    if (Array.isArray(raw)) {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        setPrompts(raw.map(function (item) {
                            return { title: item.title || '', subtitle: item.subtitle || item.description || '', icon: item.icon || '', prompt: item.prompt || item.title || '' };
                        }));
                    }
                }
                catch (_a) {
                    setPrompts([]);
                }
                loadExtendedData(result.config);
            }
            if (result.welcome && result.welcome.length > 0) {
                var sanitized = result.welcome.map(function (msg) {
                    if (msg.sender === 'bot' && msg.contentType === 'text' && typeof msg.text === 'string') {
                        return copyMessageWithText(msg, stripUsernameFromWelcome(msg.text));
                    }
                    return msg;
                });
                setMessages(sanitized);
                setShowPrompts(true);
            }
            setLoading(false);
            window.setTimeout(function () { if (textareaRef.current) {
                textareaRef.current.focus();
            } }, 100);
        }).catch(function (err) {
            setAllowed(false);
            setErrorMsg('[ERR-UNCAUGHT] ' + ((err instanceof Error) ? err.message : String(err)));
            setLoading(false);
        });
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [open]);
    // ── Feedback helpers (unchanged) ─────────────────────────────────────────────
    function hasFeedbackUrl() {
        return !!(config && config.msFormsUrl && sanitizeHttpUrl(config.msFormsUrl));
    }
    function hasMeaningfulConvo() {
        for (var i = 0; i < messages.length; i++) {
            if (messages[i].sender === 'user') {
                return true;
            }
        }
        return false;
    }
    function executePendingAction() {
        var action = pendingActionRef.current;
        pendingActionRef.current = null;
        setFeedbackStep('idle');
        if (action === 'close') {
            setClosing(true);
            window.setTimeout(function () { setOpen(false); setClosing(false); }, 200);
        }
        else if (action === 'reset') {
            if (!proxyRef.current) {
                return;
            }
            setMessages([]);
            setSending(true);
            proxyRef.current.reset().then(function (w) {
                var sanitized = w.map(function (msg) {
                    if (msg.sender === 'bot' && msg.contentType === 'text' && typeof msg.text === 'string') {
                        return copyMessageWithText(msg, stripUsernameFromWelcome(msg.text));
                    }
                    return msg;
                });
                setMessages(sanitized);
                setShowPrompts(true);
                setActiveTab('chat');
                if (config) {
                    loadExtendedData(config);
                }
                setSending(false);
                if (textareaRef.current) {
                    textareaRef.current.focus();
                }
            }).catch(function () { setSending(false); });
        }
    }
    function requestFeedback(action) {
        if (!hasFeedbackUrl() || !hasMeaningfulConvo()) {
            pendingActionRef.current = action;
            executePendingAction();
            return;
        }
        pendingActionRef.current = action;
        setFeedbackStep('rating');
    }
    function handleFeedbackRating(_liked) { setFeedbackStep('forms'); }
    function handleOpenForms() {
        if (config && config.msFormsUrl) {
            var safe = sanitizeHttpUrl(config.msFormsUrl);
            if (safe) {
                window.open(safe, '_blank', 'noopener,noreferrer');
            }
        }
        executePendingAction();
    }
    function handleFeedbackSkip() { executePendingAction(); }
    function handleOpen() { setClosing(false); setOpen(true); }
    function handleClose() {
        if (feedbackStep !== 'idle') {
            setFeedbackStep('done');
            setClosing(true);
            window.setTimeout(function () { setOpen(false); setClosing(false); setFeedbackStep('idle'); }, 200);
            return;
        }
        requestFeedback('close');
    }
    function handleReset() {
        if (feedbackStep !== 'idle') {
            setFeedbackStep('idle');
            pendingActionRef.current = 'reset';
            executePendingAction();
            return;
        }
        requestFeedback('reset');
    }
    // ── handleSend (unchanged) ──────────────────────────────────────────────────
    function handleSend(text) {
        var trimmed = text.trim();
        if (!trimmed || sending || !proxyRef.current) {
            return;
        }
        setMessages(function (prev) {
            var cleared = clearAllSuggestedActions(prev);
            var userMsg = {
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
        if (feedbackStep !== 'idle') {
            setFeedbackStep('idle');
            pendingActionRef.current = null;
        }
        setSending(true);
        if (textareaRef.current) {
            textareaRef.current.style.height = 'auto';
            textareaRef.current.focus();
        }
        proxyRef.current.send(trimmed).then(function (replies) {
            setMessages(function (prev) { return prev.concat(replies); });
            setSending(false);
            if (textareaRef.current) {
                textareaRef.current.focus();
            }
        }).catch(function () {
            setSending(false);
            if (textareaRef.current) {
                textareaRef.current.focus();
            }
        });
    }
    function handleKeyDown(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSend(inputText);
        }
    }
    function handleTextareaChange(e) {
        setInputText(e.target.value);
        var ta = e.target;
        ta.style.height = 'auto';
        ta.style.height = Math.min(ta.scrollHeight, 100) + 'px';
    }
    function handleDismissBanner(id) { persistDismiss(id); }
    // ── Derived values ───────────────────────────────────────────────────────────
    var displayName = config ? config.agentDisplayName : 'Assistant';
    var agentIconUrl = config ? config.agentIconURL : undefined;
    var panelClass = [styles.chatPanel, maximized ? styles.maximized : '', closing ? styles.closing : ''].filter(Boolean).join(' ');
    var showUpdatesTab = !!(config && config.enableNotifications && config.dataSiteUrl && config.liveUpdatesListTitle);
    var showNearbyTab = !!(config && config.enableConnectorConsent && config.dataSiteUrl && config.recommendationsListTitle);
    var showTabs = showUpdatesTab || showNearbyTab;
    var unreadCount = 0;
    var hasHighPriority = false;
    for (var i = 0; i < liveUpdates.length; i++) {
        if (!dismissed[liveUpdates[i].id]) {
            unreadCount++;
            if (liveUpdates[i].priority === 'High') {
                hasHighPriority = true;
            }
        }
    }
    /* ── Render ─────────────────────────────────────────────────────────────────── */
    return (React.createElement(React.Fragment, null,
        React.createElement("button", { className: styles.fab, onClick: open ? handleClose : handleOpen, "aria-label": "Toggle chat", title: displayName },
            React.createElement("div", { className: styles.fabInner },
                React.createElement("div", { className: styles.fabPulse }),
                React.createElement(BotAvatarIcon, { iconUrl: agentIconUrl, className: styles.fabIcon }),
                !open && showUpdatesTab && unreadCount > 0 ? (React.createElement("div", { className: styles.fabNotifBadge, "aria-label": unreadCount + ' unread updates' }, unreadCount > 9 ? '9+' : unreadCount)) : null),
            !open && labelText ? (React.createElement("div", { className: styles.fabLabel + (showLabel ? ' ' + styles.visible : '') }, labelText)) : null),
        open ? (React.createElement("div", { className: panelClass, role: "dialog", "aria-label": displayName + ' chat' },
            React.createElement("div", { className: styles.chatHeader },
                React.createElement("div", { className: styles.headerLeft },
                    React.createElement("div", { className: styles.headerAvatar },
                        React.createElement(BotAvatarIcon, { iconUrl: agentIconUrl, className: styles.headerAvatarIcon })),
                    React.createElement("div", { className: styles.headerInfo },
                        React.createElement("div", { className: styles.headerTitle }, displayName),
                        React.createElement("div", { className: styles.headerSubtitle }, loading ? 'Connecting…' : (allowed ? 'Powered by Copilot Studio' : '')))),
                React.createElement("div", { className: styles.headerActions },
                    React.createElement("button", { className: styles.headerBtn, onClick: handleReset, title: "New conversation", disabled: loading || !allowed },
                        React.createElement(ArrowCounterclockwise20Regular, null)),
                    React.createElement("button", { className: styles.headerBtn, onClick: function () { setMaximized(function (m) { return !m; }); }, title: maximized ? 'Restore' : 'Expand' }, maximized ? React.createElement(ArrowMinimize20Regular, null) : React.createElement(ArrowMaximize20Regular, null)),
                    React.createElement("button", { className: styles.headerBtnClose, onClick: handleClose, title: "Close" },
                        React.createElement(Dismiss20Regular, null)))),
            showTabs && !loading && allowed ? (React.createElement("div", { className: styles.tabBar, role: "tablist" },
                React.createElement("button", { role: "tab", "aria-selected": activeTab === 'chat', className: styles.tabBtn + (activeTab === 'chat' ? ' ' + styles.tabBtnActive : ''), onClick: function () { setActiveTab('chat'); } },
                    React.createElement(Chat20Regular, { className: styles.tabIcon }),
                    React.createElement("span", null, "Chat")),
                showUpdatesTab ? (React.createElement("button", { role: "tab", "aria-selected": activeTab === 'updates', className: styles.tabBtn + (activeTab === 'updates' ? ' ' + styles.tabBtnActive : '') + (hasHighPriority ? ' ' + styles.tabBtnAlert : ''), onClick: function () {
                        setActiveTab('updates');
                        for (var j = 0; j < liveUpdates.length; j++) {
                            persistDismiss(liveUpdates[j].id);
                        }
                    } },
                    React.createElement(Alert20Regular, { className: styles.tabIcon }),
                    React.createElement("span", null, "Updates"),
                    unreadCount > 0 ? React.createElement("span", { className: styles.tabBadge }, unreadCount) : null)) : null,
                showNearbyTab ? (React.createElement("button", { role: "tab", "aria-selected": activeTab === 'nearby', className: styles.tabBtn + (activeTab === 'nearby' ? ' ' + styles.tabBtnActive : ''), onClick: function () { setActiveTab('nearby'); } },
                    React.createElement(Map20Regular, { className: styles.tabIcon }),
                    React.createElement("span", null, "Nearby"))) : null)) : null,
            loading ? (React.createElement("div", { className: styles.loadingState },
                React.createElement("div", { className: styles.loadingDot }),
                React.createElement("div", { className: styles.loadingDot }),
                React.createElement("div", { className: styles.loadingDot }))) : !allowed ? (React.createElement("div", { className: styles.blockedState },
                React.createElement(LockClosed24Regular, { className: styles.blockedIcon }),
                React.createElement("div", { className: styles.blockedTitle }, errorMsg.indexOf('[ERR-CONFIG]') !== -1 ? 'Configuration not found' : errorMsg.indexOf('[ERR-DISABLED]') !== -1 ? 'Chatbot disabled' : errorMsg.indexOf('[ERR-RBAC]') !== -1 ? 'Access restricted' : errorMsg.indexOf('[ERR-MSAL]') !== -1 ? 'Authentication error' : errorMsg.indexOf('[ERR-LIST-THROW]') !== -1 ? 'List connection error' : errorMsg.indexOf('[ERR-UNCAUGHT]') !== -1 ? 'Unexpected error' : 'Access restricted'),
                React.createElement("div", { className: styles.blockedDesc }, errorMsg))) : (React.createElement(React.Fragment, null,
                activeTab === 'updates' && showUpdatesTab ? React.createElement(UpdatesTab, { updates: liveUpdates }) : null,
                activeTab === 'nearby' && showNearbyTab ? (React.createElement(NearbyTab, { recommendations: recommendations, geocodeContext: config ? (config.geocodeContext || '') : '' })) : null,
                activeTab === 'chat' ? (React.createElement(React.Fragment, null,
                    showUpdatesTab ? React.createElement(LiveUpdateBanner, { updates: liveUpdates, dismissed: dismissed, onDismiss: handleDismissBanner }) : null,
                    showPrompts && messages.length <= 2 && prompts.length > 0 ? (React.createElement("div", { className: styles.promptsSection },
                        React.createElement("div", { className: styles.promptsLabel },
                            React.createElement(Sparkle20Regular, { className: styles.promptsLabelIcon }),
                            React.createElement("span", null, "Suggested prompts")),
                        React.createElement("div", { className: styles.suggestedPrompts }, prompts.slice(0, 6).map(function (p, i) {
                            return (React.createElement("button", { key: i, className: styles.promptChip, onClick: function () { handleSend(p.prompt); } },
                                React.createElement("div", { className: styles.promptIconWrap }, getPromptIcon(p.icon || p.title)),
                                React.createElement("div", { className: styles.promptText },
                                    React.createElement("div", { className: styles.promptTitle }, p.title))));
                        })))) : null,
                    React.createElement("div", { className: styles.messagesContainer },
                        messages.map(function (msg) {
                            var isUser = msg.sender === 'user';
                            var rowClass = styles.messageRow + (isUser ? ' ' + styles.userRow : '');
                            var bubClass = styles.bubble + ' ' + (isUser ? styles.userBubble : msg.sender === 'system' ? styles.errorBubble : styles.botBubble);
                            return (React.createElement("div", { key: msg.id, ref: function (el) { messageRefs.current[msg.id] = el; }, className: rowClass },
                                msg.sender === 'bot' ? React.createElement("div", { className: styles.msgAvatar },
                                    React.createElement(BotAvatarIcon, { iconUrl: agentIconUrl, className: styles.msgAvatarIcon })) : null,
                                React.createElement("div", { className: styles.msgContent },
                                    msg.contentType === 'adaptiveCard' && msg.cardPayload ? (React.createElement("div", { className: styles.bubble + ' ' + styles.botBubble },
                                        React.createElement(AdaptiveCardHost, { payload: msg.cardPayload }))) : (React.createElement("div", { className: bubClass }, (function () {
                                        var ex = extractCitationsFromText(msg.text || '');
                                        var byIdx = {};
                                        for (var ci = 0; ci < ex.citations.length; ci++) {
                                            byIdx[ex.citations[ci].index] = ex.citations[ci];
                                        }
                                        var onCC = function (n) { toggleRefs(msg.id, true); var el = messageRefs.current[msg.id] || null; if (el) {
                                            el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                                        } };
                                        return (React.createElement(React.Fragment, null,
                                            msg.sender === 'bot' && msg.text ? renderMarkdown(ex.cleanText || msg.text, { citationsByIndex: byIdx, onCitationClick: onCC }) : msg.text,
                                            ex.citations.length > 0 ? (React.createElement("div", { className: styles.refsSection }, (function () {
                                                var isOpen = refsOpen[msg.id] !== undefined ? refsOpen[msg.id] : true;
                                                return (React.createElement(React.Fragment, null,
                                                    React.createElement("button", { type: "button", className: styles.refsToggle, onClick: function () { toggleRefs(msg.id); }, "aria-expanded": isOpen },
                                                        ex.citations.length,
                                                        " reference",
                                                        ex.citations.length > 1 ? 's' : '',
                                                        ' ',
                                                        React.createElement("span", { className: styles.chev }, refsOpen[msg.id] ? '▾' : '▸')),
                                                    refsOpen[msg.id] ? (React.createElement("div", { className: styles.refsList }, ex.citations.slice().sort(function (a, b) { return a.index - b.index; }).map(function (c) {
                                                        return (React.createElement("a", { key: c.index, href: sanitizeHttpUrl(c.url) || '#', target: "_blank", rel: "noopener noreferrer", className: styles.refChip },
                                                            React.createElement("span", { className: styles.refIndex }, c.index),
                                                            React.createElement("span", { className: styles.refLabel }, niceLabelFromUrl(c.url, c.title))));
                                                    }))) : null));
                                            })())) : null));
                                    })())),
                                    msg.sender === 'bot' && msg.suggestedActions && msg.suggestedActions.length > 0 ? (React.createElement("div", { className: styles.suggestedActionsRow }, msg.suggestedActions.map(function (action, ai) {
                                        return (React.createElement("button", { key: ai, className: styles.suggestedActionChip, onClick: function () { handleSend(action.value); }, disabled: sending, title: action.title }, action.title));
                                    }))) : null,
                                    React.createElement("div", { className: styles.msgTime + (isUser ? ' ' + styles.msgTimeUser : '') }, formatTime(msg.timestamp)))));
                        }),
                        sending ? (React.createElement("div", { className: styles.typingRow },
                            React.createElement("div", { className: styles.msgAvatar },
                                React.createElement(BotAvatarIcon, { iconUrl: agentIconUrl, className: styles.msgAvatarIcon })),
                            React.createElement("div", { className: styles.typingBubble },
                                React.createElement("span", null),
                                React.createElement("span", null),
                                React.createElement("span", null)))) : null,
                        React.createElement("div", { ref: messagesEndRef })),
                    React.createElement(FeedbackDialog, { step: feedbackStep, onLike: function () { handleFeedbackRating(true); }, onDislike: function () { handleFeedbackRating(false); }, onOpenForms: handleOpenForms, onSkip: handleFeedbackSkip }),
                    React.createElement("div", { className: styles.inputArea },
                        React.createElement("div", { className: styles.inputWrap },
                            React.createElement("textarea", { ref: textareaRef, className: styles.textInput, placeholder: "Ask a question\u2026", value: inputText, onChange: handleTextareaChange, onKeyDown: handleKeyDown, rows: 1, disabled: sending }),
                            React.createElement("button", { className: styles.sendBtn, onClick: function () { handleSend(inputText); }, disabled: !inputText.trim() || sending, title: "Send" },
                                React.createElement(Send20Filled, null))),
                        React.createElement("div", { className: styles.inputFooter }, "Powered by Microsoft Copilot Studio")))) : null)))) : null));
};
//# sourceMappingURL=ChatWidget.js.map