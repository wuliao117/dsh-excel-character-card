/**
 * dsh-excel-character-card — browser half.
 *
 * A compact "角色卡" button in the composer toolbar opens a panel where the
 * user can drag in an .xlsx/.xlsm character sheet, preview the parsed cards,
 * activate one for the current session, or stop role-play.  All storage and
 * parsing stay on the Host; this file only talks to the plugin's JSON API.
 */
window.__ModuleLoader__.load({
  id: "dsh-excel-character-card",
  factory: function (require) {
    var module = { exports: {} };
    var exports = module.exports;
    var React = require("react");
    var ReactDOM = require("react-dom");
    var h = React.createElement;

    var API = "/excel-character-card/api";
    var PACKAGE_ID = "dsh-excel-character-card";
    var ICON_CARD =
      '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<rect x="3" y="4" width="18" height="16" rx="2.5"></rect>' +
      '<path d="M7.5 8.5h9M7.5 12h6M7.5 15.5h7.5"></path></svg>';
    var ICON_CLOSE =
      '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true">' +
      '<path d="M6 6l12 12M18 6L6 18"></path></svg>';
    var ICON_BACK =
      '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M15 5l-7 7 7 7"></path></svg>';

    var FIELD_GROUPS = [
      {
        title: "基础",
        fields: [
          ["gender", "性别"],
          ["age", "年龄"],
          ["birthday", "生日"],
          ["height", "身高"],
          ["weight", "体重"],
          ["species", "种族"],
          ["identity", "身份 / 职业"],
        ],
      },
      { title: "外貌", fields: [["appearance", "外貌"]] },
      {
        title: "性格",
        fields: [
          ["personality", "性格"],
          ["likes", "喜好"],
          ["dislikes", "讨厌"],
          ["fears", "恐惧"],
        ],
      },
      {
        title: "背景",
        fields: [
          ["background", "背景 / 经历"],
          ["motivation", "动机 / 目标"],
          ["scenario", "场景 / 世界观"],
        ],
      },
      {
        title: "能力",
        fields: [
          ["skills", "技能 / 特长"],
          ["weaknesses", "弱点"],
        ],
      },
      {
        title: "语言",
        fields: [
          ["speechStyle", "说话风格"],
          ["catchphrases", "口头禅"],
          ["quote", "名言 / 座右铭"],
          ["greeting", "开场白"],
        ],
      },
      { title: "关系", fields: [["relationships", "人际关系"]] },
      { title: "秘密", fields: [["secrets", "秘密"]] },
      { title: "规则", fields: [["rules", "规则 / 禁忌"]] },
    ];

    var CSS = [
      ".ecc-trigger{display:inline-flex;align-items:center;gap:6px;height:28px;padding:0 9px;border:1px solid var(--dsw-alias-border-l1,rgba(15,17,21,.12));border-radius:8px;background:transparent;color:var(--dsw-alias-label-secondary,#5c6470);font-size:12px;line-height:1;cursor:pointer;white-space:nowrap;transition:background .15s,color .15s,border-color .15s}",
      ".ecc-trigger:hover{background:var(--dsw-alias-bg-layer-2,#f1f3f5);color:var(--dsw-alias-label-primary,#0f1115)}",
      ".ecc-trigger:focus-visible{outline:2px solid #4d93f8;outline-offset:2px}",
      ".ecc-trigger svg{flex:none;opacity:.9}",
      ".ecc-overlay{position:fixed;inset:0;z-index:2147483000;display:flex;align-items:flex-start;justify-content:flex-end;padding:72px 24px 24px;background:rgba(8,10,14,.38);backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px)}",
      ".ecc-panel{--ecc-bg:var(--dsw-alias-bg-overlay,#fff);--ecc-soft:var(--dsw-alias-bg-layer-1,#f7f8fa);--ecc-text:var(--dsw-alias-label-primary,#0f1115);--ecc-muted:var(--dsw-alias-label-secondary,#6b7280);--ecc-border:var(--dsw-alias-border-l1,rgba(15,17,21,.1));--ecc-accent:#4d93f8;--ecc-danger:var(--dsw-alias-state-error-primary,#e5484d);width:440px;max-width:calc(100vw - 32px);max-height:calc(100vh - 96px);display:flex;flex-direction:column;overflow:hidden;background:var(--ecc-bg);color:var(--ecc-text);border:1px solid var(--dsw-alias-border-l2,rgba(15,17,21,.14));border-radius:16px;box-shadow:0 24px 80px rgba(0,0,0,.26);font-family:var(--dsw-font-family,system-ui,-apple-system,'Segoe UI',sans-serif)}",
      "body[data-ds-dark-theme] .ecc-panel{--ecc-bg:var(--dsw-alias-bg-overlay,#1b1c20);--ecc-soft:var(--dsw-alias-bg-layer-1,#24252a);--ecc-text:var(--dsw-alias-label-primary,#f2f3f5);--ecc-muted:var(--dsw-alias-label-secondary,#a5a9b3);--ecc-border:var(--dsw-alias-border-l1,rgba(255,255,255,.12))}",
      ".ecc-header{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;padding:16px 18px 14px;border-bottom:1px solid var(--ecc-border)}",
      ".ecc-title-wrap{display:flex;align-items:center;gap:10px;min-width:0}",
      ".ecc-logo{width:34px;height:34px;flex:none;display:flex;align-items:center;justify-content:center;border-radius:10px;background:color-mix(in srgb,var(--ecc-accent) 14%,transparent);color:var(--ecc-accent)}",
      ".ecc-title{font-size:14px;font-weight:650;line-height:1.35}",
      ".ecc-subtitle{margin-top:2px;font-size:11px;line-height:1.35;color:var(--ecc-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:280px}",
      ".ecc-icon-btn{width:30px;height:30px;flex:none;display:flex;align-items:center;justify-content:center;border:0;border-radius:8px;background:transparent;color:var(--ecc-muted);cursor:pointer}",
      ".ecc-icon-btn:hover{background:var(--ecc-soft);color:var(--ecc-text)}",
      ".ecc-scroll{overflow-y:auto;overscroll-behavior:contain;padding:16px 18px;display:flex;flex-direction:column;gap:14px}",
      ".ecc-footer{padding:10px 18px;border-top:1px solid var(--ecc-border);font-size:10.5px;line-height:1.45;color:var(--ecc-muted);display:flex;align-items:center;justify-content:space-between;gap:10px}",
      ".ecc-footer code{font-size:10px;background:var(--ecc-soft);padding:1px 4px;border-radius:4px}",
      ".ecc-drop{display:flex;flex-direction:column;align-items:center;gap:5px;padding:18px 14px;border:1px dashed var(--ecc-border);border-radius:12px;background:var(--ecc-soft);cursor:pointer;text-align:center;transition:border-color .15s,background .15s}",
      ".ecc-drop:hover,.ecc-drop.is-dragging{border-color:var(--ecc-accent);background:color-mix(in srgb,var(--ecc-accent) 7%,var(--ecc-soft))}",
      ".ecc-drop-title{font-size:12.5px;font-weight:600}",
      ".ecc-drop-sub{font-size:11px;color:var(--ecc-muted)}",
      ".ecc-drop svg{opacity:.75;color:var(--ecc-accent)}",
      ".ecc-active{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:11px 12px;border:1px solid color-mix(in srgb,var(--ecc-accent) 30%,var(--ecc-border));border-radius:11px;background:color-mix(in srgb,var(--ecc-accent) 8%,transparent)}",
      ".ecc-active-main{min-width:0}",
      ".ecc-active-label{font-size:10px;letter-spacing:.04em;text-transform:uppercase;color:var(--ecc-accent);font-weight:700}",
      ".ecc-active-name{font-size:13px;font-weight:650;margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}",
      ".ecc-section-head{display:flex;align-items:center;justify-content:space-between;gap:10px}",
      ".ecc-section-title{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--ecc-muted)}",
      ".ecc-list{display:flex;flex-direction:column;gap:8px}",
      ".ecc-card{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;padding:11px 12px;border:1px solid var(--ecc-border);border-radius:11px;background:var(--ecc-bg)}",
      ".ecc-card.is-active{border-color:color-mix(in srgb,var(--ecc-accent) 45%,var(--ecc-border))}",
      ".ecc-card-main{min-width:0;flex:1}",
      ".ecc-card-name{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:650;line-height:1.35}",
      ".ecc-card-tagline{margin-top:2px;font-size:11px;line-height:1.4;color:var(--ecc-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}",
      ".ecc-card-source{margin-top:3px;font-size:10px;line-height:1.35;color:var(--ecc-muted);opacity:.8;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}",
      ".ecc-card-actions{display:flex;align-items:center;gap:5px;flex:none}",
      ".ecc-badge{display:inline-flex;align-items:center;height:17px;padding:0 5px;border-radius:5px;background:color-mix(in srgb,var(--ecc-accent) 14%,transparent);color:var(--ecc-accent);font-size:9.5px;font-weight:700;letter-spacing:.03em}",
      ".ecc-btn{height:26px;padding:0 8px;border:1px solid var(--ecc-border);border-radius:7px;background:transparent;color:var(--ecc-text);font-size:11px;line-height:1;cursor:pointer;white-space:nowrap}",
      ".ecc-btn:hover:not(:disabled){background:var(--ecc-soft)}",
      ".ecc-btn:disabled{opacity:.55;cursor:default}",
      ".ecc-btn-primary{border-color:transparent;background:var(--ecc-accent);color:#fff}",
      ".ecc-btn-primary:hover:not(:disabled){filter:brightness(1.05);background:var(--ecc-accent)}",
      ".ecc-btn-danger{color:var(--ecc-danger)}",
      ".ecc-btn-danger:hover:not(:disabled){background:color-mix(in srgb,var(--ecc-danger) 8%,transparent)}",
      ".ecc-empty{padding:24px 12px;text-align:center;border:1px dashed var(--ecc-border);border-radius:12px;color:var(--ecc-muted);font-size:11.5px;line-height:1.6}",
      ".ecc-error,.ecc-notice{padding:8px 10px;border-radius:8px;font-size:11.5px;line-height:1.45}",
      ".ecc-error{color:var(--ecc-danger);background:color-mix(in srgb,var(--ecc-danger) 8%,transparent);border:1px solid color-mix(in srgb,var(--ecc-danger) 24%,transparent)}",
      ".ecc-notice{color:var(--ecc-accent);background:color-mix(in srgb,var(--ecc-accent) 8%,transparent);border:1px solid color-mix(in srgb,var(--ecc-accent) 22%,transparent)}",
      ".ecc-detail-head{display:flex;align-items:center;gap:8px;min-width:0}",
      ".ecc-detail-name{font-size:16px;font-weight:700;line-height:1.3}",
      ".ecc-detail-tagline{font-size:11.5px;color:var(--ecc-muted);line-height:1.45}",
      ".ecc-group{border-top:1px solid var(--ecc-border);padding-top:10px}",
      ".ecc-group:first-of-type{border-top:0;padding-top:0}",
      ".ecc-group-title{font-size:10.5px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--ecc-muted);margin-bottom:6px}",
      ".ecc-field{display:flex;gap:10px;padding:4px 0;font-size:11.5px;line-height:1.55}",
      ".ecc-field-label{flex:none;width:82px;color:var(--ecc-muted)}",
      ".ecc-field-value{flex:1;min-width:0;white-space:pre-wrap;word-break:break-word}",
      ".ecc-prompt{max-height:220px;overflow:auto;margin:0;padding:10px;border:1px solid var(--ecc-border);border-radius:8px;background:var(--ecc-soft);font-size:10.5px;line-height:1.55;white-space:pre-wrap;word-break:break-word;font-family:ui-monospace,SFMono-Regular,Consolas,monospace}",
      ".ecc-details summary{cursor:pointer;font-size:11px;color:var(--ecc-muted);padding:4px 0}",
      ".ecc-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}",
      ".ecc-spacer{flex:1}",
      "@media (max-width:640px){.ecc-overlay{padding:0;align-items:flex-end}.ecc-panel{width:100%;max-width:100%;max-height:88vh;border-radius:16px 16px 0 0}.ecc-trigger-label{display:none}.ecc-subtitle{max-width:210px}.ecc-field-label{width:70px}}",
      "@media (prefers-reduced-motion:reduce){.ecc-trigger,.ecc-drop{transition:none}}",
    ].join("\n");

    function installStyles() {
      try {
        if (typeof document === "undefined" || !document.head) return;
        if (document.querySelector('style[data-plugin-css="' + PACKAGE_ID + '"]')) return;
        var style = document.createElement("style");
        style.setAttribute("data-plugin-css", PACKAGE_ID);
        style.textContent = CSS;
        document.head.appendChild(style);
      } catch (_) {
        /* styling failure must not break the panel */
      }
    }

    function api(path, options) {
      var init = options || {};
      var headers = { accept: "application/json" };
      if (init.body !== undefined) headers["content-type"] = "application/json";
      return fetch(API + path, {
        method: init.method || "GET",
        headers: headers,
        credentials: "same-origin",
        body: init.body,
      }).then(function (response) {
        return response
          .json()
          .catch(function () {
            return null;
          })
          .then(function (data) {
            if (!response.ok || data === null || data.ok === false) {
              var message =
                response.status === 404 && (!data || typeof data.error !== "string")
                  ? "角色卡宿主插件尚未加载：请完全退出并重新打开 DSH 后重试。"
                  : data && typeof data.error === "string"
                    ? data.error
                    : "HTTP " + response.status;
              throw new Error(message);
            }
            return data.result;
          });
      });
    }

    function fileToBase64(file) {
      return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onerror = function () {
          reject(new Error("读取文件失败"));
        };
        reader.onload = function () {
          var result = String(reader.result || "");
          var comma = result.indexOf(",");
          resolve(comma >= 0 ? result.slice(comma + 1) : result);
        };
        reader.readAsDataURL(file);
      });
    }

    function errorText(error) {
      if (error && typeof error.message === "string" && error.message !== "") return error.message;
      return String(error);
    }

    function Svg(props) {
      return h("span", {
        style: { display: "inline-flex", alignItems: "center", justifyContent: "center" },
        dangerouslySetInnerHTML: { __html: props.html },
      });
    }

    function FieldGroup(props) {
      var card = props.card;
      var fields = props.fields.filter(function (pair) {
        return typeof card.fields[pair[0]] === "string" && card.fields[pair[0]].trim() !== "";
      });
      if (fields.length === 0) return null;
      return h(
        "div",
        { className: "ecc-group" },
        h("div", { className: "ecc-group-title" }, props.title),
        fields.map(function (pair) {
          return h(
            "div",
            { className: "ecc-field", key: pair[0] },
            h("div", { className: "ecc-field-label" }, pair[1]),
            h("div", { className: "ecc-field-value" }, card.fields[pair[0]]),
          );
        }),
      );
    }

    function DetailView(props) {
      var card = props.card;
      var busy = props.busy;
      if (card === null) {
        return h(
          "div",
          null,
          h(
            "div",
            { className: "ecc-row", style: { marginBottom: "10px" } },
            h(
              "button",
              { className: "ecc-icon-btn", type: "button", onClick: props.onBack, title: "返回" },
              h(Svg, { html: ICON_BACK }),
            ),
            h("div", { className: "ecc-detail-name" }, props.error ? "读取失败" : "正在读取角色卡…"),
          ),
          props.error ? h("div", { className: "ecc-error" }, props.error) : null,
        );
      }
      var isActive = props.activeCardId === card.id;
      return h(
        "div",
        { style: { display: "flex", flexDirection: "column", gap: "12px" } },
        h(
          "div",
          { className: "ecc-row" },
          h(
            "button",
            { className: "ecc-icon-btn", type: "button", onClick: props.onBack, title: "返回" },
            h(Svg, { html: ICON_BACK }),
          ),
          h("div", { className: "ecc-spacer" }),
          isActive
            ? h("button", { className: "ecc-btn", type: "button", disabled: busy, onClick: props.onDeactivate }, "停用")
            : h(
                "button",
                { className: "ecc-btn ecc-btn-primary", type: "button", disabled: busy, onClick: props.onActivate },
                "启用",
              ),
        ),
        h(
          "div",
          null,
          h("div", { className: "ecc-detail-name" }, card.name, isActive ? h("span", { className: "ecc-badge", style: { marginLeft: "8px" } }, "当前") : null),
          h("div", { className: "ecc-detail-tagline" }, [card.tagline, card.source.fileName + " / " + card.source.sheet].filter(Boolean).join(" · ")),
        ),
        FIELD_GROUPS.map(function (group) {
          return h(FieldGroup, { key: group.title, title: group.title, fields: group.fields, card: card });
        }),
        card.extra && card.extra.length > 0
          ? h(
              "div",
              { className: "ecc-group" },
              h("div", { className: "ecc-group-title" }, "其他设定"),
              card.extra.map(function (item, index) {
                return h(
                  "div",
                  { className: "ecc-field", key: index },
                  h("div", { className: "ecc-field-label" }, item.label),
                  h("div", { className: "ecc-field-value" }, item.value),
                );
              }),
            )
          : null,
        card.prompt
          ? h(
              "details",
              { className: "ecc-details" },
              h("summary", null, "查看注入的系统提示词"),
              h("pre", { className: "ecc-prompt" }, card.prompt),
            )
          : null,
      );
    }

    function Panel(props) {
      var sessionId = props.sessionId;
      var [cards, setCards] = React.useState([]);
      var [activeCardId, setActiveCardId] = React.useState(null);
      var [activeCard, setActiveCard] = React.useState(null);
      var [activeScope, setActiveScope] = React.useState("none");
      var [loading, setLoading] = React.useState(true);
      var [busy, setBusy] = React.useState("");
      var [error, setError] = React.useState("");
      var [notice, setNotice] = React.useState("");
      var [detail, setDetail] = React.useState(null);
      var [detailCardId, setDetailCardId] = React.useState(null);
      var [dragging, setDragging] = React.useState(false);
      var fileRef = React.useRef(null);

      function applyState(result) {
        if (!result) return;
        setCards(Array.isArray(result.cards) ? result.cards : []);
        setActiveCardId(typeof result.activeCardId === "string" ? result.activeCardId : null);
        setActiveCard(result.activeCard || null);
        setActiveScope(typeof result.activeScope === "string" ? result.activeScope : "none");
      }

      function load() {
        setLoading(true);
        setError("");
        api("/state?sessionId=" + encodeURIComponent(sessionId || ""))
          .then(function (result) {
            applyState(result);
          })
          .catch(function (err) {
            setError(errorText(err));
          })
          .finally(function () {
            setLoading(false);
          });
      }

      React.useEffect(function () {
        load();
        // eslint-disable-next-line react-hooks/exhaustive-deps
      }, [sessionId]);

      React.useEffect(function () {
        function onKey(event) {
          if (event.key === "Escape") props.onClose();
        }
        window.addEventListener("keydown", onKey);
        return function () {
          window.removeEventListener("keydown", onKey);
        };
      }, [props.onClose]);

      function run(key, operation) {
        setBusy(key);
        setError("");
        setNotice("");
        return operation()
          .catch(function (err) {
            setError(errorText(err));
          })
          .finally(function () {
            setBusy("");
          });
      }

      function upload(file) {
        if (!file) return;
        if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
          setError("请选择 .xlsx 或 .xlsm 文件（旧版 .xls 请先另存为 .xlsx）。");
          return;
        }
        return run("upload", async function () {
          var base64 = await fileToBase64(file);
          var result = await api("/import", {
            method: "POST",
            body: JSON.stringify({ sessionId: sessionId, fileName: file.name, dataBase64: base64 }),
          });
          applyState(result);
          setNotice("已导入 " + file.name + "（新增 " + result.report.added + "，更新 " + result.report.updated + "）。");
        });
      }

      function activate(cardId, name) {
        return run("activate:" + cardId, async function () {
          var result = await api("/activate", {
            method: "POST",
            body: JSON.stringify({ sessionId: sessionId, cardId: cardId }),
          });
          applyState(result);
          setNotice("已启用「" + (name || "角色") + "」，之后会以该角色的身份回复。");
        });
      }

      function deactivate() {
        return run("deactivate", async function () {
          var result = await api("/deactivate", {
            method: "POST",
            body: JSON.stringify({ sessionId: sessionId }),
          });
          applyState(result);
          setNotice("已停止角色扮演。");
        });
      }

      function remove(card) {
        if (!window.confirm("删除角色卡「" + card.name + "」？此操作不可撤销。")) return;
        return run("delete:" + card.id, async function () {
          var result = await api("/delete", {
            method: "POST",
            body: JSON.stringify({ sessionId: sessionId, cardId: card.id }),
          });
          applyState(result);
          if (detailCardId === card.id) {
            setDetail(null);
            setDetailCardId(null);
          }
          setNotice("已删除「" + card.name + "」。");
        });
      }

      function openDetail(cardId) {
        setDetailCardId(cardId);
        setDetail(null);
        setError("");
        api("/card?id=" + encodeURIComponent(cardId))
          .then(function (result) {
            setDetail(Object.assign({}, result.card, { prompt: result.prompt }));
          })
          .catch(function (err) {
            setError(errorText(err));
          });
      }

      function closeDetail() {
        setDetail(null);
        setDetailCardId(null);
      }

      var content;
      if (detailCardId !== null) {
        content = h(DetailView, {
          card: detail,
          error: error,
          activeCardId: activeCardId,
          busy: busy !== "",
          onBack: closeDetail,
          onActivate: function () {
            activate(detailCardId, detail && detail.name);
          },
          onDeactivate: deactivate,
        });
      } else {
        content = h(
          React.Fragment,
          null,
          h(
            "label",
            {
              className: "ecc-drop" + (dragging ? " is-dragging" : ""),
              onDragOver: function (event) {
                event.preventDefault();
                setDragging(true);
              },
              onDragLeave: function () {
                setDragging(false);
              },
              onDrop: function (event) {
                event.preventDefault();
                setDragging(false);
                upload(event.dataTransfer.files && event.dataTransfer.files[0]);
              },
            },
            h("input", {
              ref: fileRef,
              type: "file",
              accept: ".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12",
              style: { display: "none" },
              onChange: function (event) {
                var file = event.target.files && event.target.files[0];
                event.target.value = "";
                upload(file);
              },
            }),
            h(Svg, { html: ICON_CARD }),
            h("div", { className: "ecc-drop-title" }, busy === "upload" ? "正在解析 Excel…" : "导入 Excel 角色卡"),
            h("div", { className: "ecc-drop-sub" }, "拖入 .xlsx / .xlsm，或点击选择文件"),
          ),
          error ? h("div", { className: "ecc-error" }, error) : null,
          notice ? h("div", { className: "ecc-notice" }, notice) : null,
          activeCard
            ? h(
                "div",
                { className: "ecc-active" },
                h(
                  "div",
                  { className: "ecc-active-main" },
                  h("div", { className: "ecc-active-label" }, activeScope === "global" ? "全局默认角色" : "当前会话角色"),
                  h("div", { className: "ecc-active-name" }, activeCard.name),
                  activeCard.tagline ? h("div", { className: "ecc-card-tagline" }, activeCard.tagline) : null,
                ),
                h(
                  "button",
                  { className: "ecc-btn", type: "button", disabled: busy !== "", onClick: deactivate },
                  "停用",
                ),
              )
            : null,
          h(
            "div",
            { className: "ecc-section-head" },
            h("div", { className: "ecc-section-title" }, "角色卡" + (cards.length > 0 ? " · " + cards.length : "")),
            h(
              "button",
              { className: "ecc-btn", type: "button", disabled: loading || busy !== "", onClick: load },
              loading ? "读取中…" : "刷新",
            ),
          ),
          loading
            ? h("div", { className: "ecc-empty" }, "正在读取角色卡…")
            : cards.length === 0
              ? h("div", { className: "ecc-empty" }, "还没有角色卡。\n上传一个 Excel：A 列写字段名（姓名、性别、年龄、身高、性格…），B 列写内容；或者第一行写表头、下面每行一个角色。")
              : h(
                  "div",
                  { className: "ecc-list" },
                  cards.map(function (card) {
                    var isActive = activeCardId === card.id;
                    return h(
                      "div",
                      { className: "ecc-card" + (isActive ? " is-active" : ""), key: card.id },
                      h(
                        "div",
                        { className: "ecc-card-main" },
                        h(
                          "div",
                          { className: "ecc-card-name" },
                          card.name,
                          isActive ? h("span", { className: "ecc-badge" }, "当前") : null,
                        ),
                        card.tagline ? h("div", { className: "ecc-card-tagline" }, card.tagline) : null,
                        h("div", { className: "ecc-card-source" }, card.source.fileName + " / " + card.source.sheet + " · " + card.fieldCount + " 项设定"),
                      ),
                      h(
                        "div",
                        { className: "ecc-card-actions" },
                        isActive
                          ? h("button", { className: "ecc-btn", type: "button", disabled: busy !== "", onClick: deactivate }, "停用")
                          : h(
                              "button",
                              {
                                className: "ecc-btn ecc-btn-primary",
                                type: "button",
                                disabled: busy !== "",
                                onClick: function () {
                                  activate(card.id, card.name);
                                },
                              },
                              "启用",
                            ),
                        h(
                          "button",
                          {
                            className: "ecc-btn",
                            type: "button",
                            onClick: function () {
                              openDetail(card.id);
                            },
                          },
                          "详情",
                        ),
                        h(
                          "button",
                          {
                            className: "ecc-btn ecc-btn-danger",
                            type: "button",
                            disabled: busy !== "",
                            onClick: function () {
                              remove(card);
                            },
                          },
                          "删除",
                        ),
                      ),
                    );
                  }),
                ),
        );
      }

      return ReactDOM.createPortal(
        h(
          "div",
          {
            className: "ecc-overlay",
            onMouseDown: function (event) {
              if (event.target === event.currentTarget) props.onClose();
            },
          },
          h(
            "div",
            { className: "ecc-panel", role: "dialog", "aria-modal": "true", "aria-label": "Excel 角色卡" },
            h(
              "div",
              { className: "ecc-header" },
              h(
                "div",
                { className: "ecc-title-wrap" },
                h("div", { className: "ecc-logo" }, h(Svg, { html: ICON_CARD })),
                h(
                  "div",
                  null,
                  h("div", { className: "ecc-title" }, "Excel 角色卡"),
                  h(
                    "div",
                    { className: "ecc-subtitle" },
                    activeCard ? "当前扮演：" + activeCard.name : "上传角色表，让助手进入角色",
                  ),
                ),
              ),
              h(
                "button",
                { className: "ecc-icon-btn", type: "button", onClick: props.onClose, title: "关闭" },
                h(Svg, { html: ICON_CLOSE }),
              ),
            ),
            h("div", { className: "ecc-scroll" }, content),
            h(
              "div",
              { className: "ecc-footer" },
              h("span", null, "数据仅保存在本机 ", h("code", null, "~/.dsh/excel-character-card")),
              detailCardId === null
                ? h(
                    "button",
                    {
                      className: "ecc-btn",
                      type: "button",
                      onClick: function () {
                        if (fileRef.current) fileRef.current.click();
                      },
                      disabled: busy !== "",
                    },
                    "选择文件",
                  )
                : h("span", null, "返回列表可导入新文件"),
            ),
          ),
        ),
        document.body,
      );
    }

    function Trigger(props) {
      var [open, setOpen] = React.useState(false);
      return h(
        React.Fragment,
        null,
        h(
          "button",
          {
            type: "button",
            className: "ecc-trigger",
            title: "Excel 角色卡：导入角色表并进入角色扮演",
            "aria-haspopup": "dialog",
            onClick: function () {
              setOpen(true);
            },
          },
          h(Svg, { html: ICON_CARD }),
          h("span", { className: "ecc-trigger-label" }, "角色卡"),
        ),
        open
          ? h(Panel, {
              sessionId: props.sessionId,
              onClose: function () {
                setOpen(false);
              },
            })
          : null,
      );
    }

    function apply(ctx) {
      installStyles();
      try {
        ctx.slots.inject("conversation.input.right", function () {
          return ctx.slots.register(
            {
              name: "conversation.input.right",
              id: "excel-character-card",
              order: 55,
            },
            Trigger,
          );
        });
      } catch (error) {
        try {
          console.warn("[excel-character-card] 无法挂载输入框按钮：", error);
        } catch (_) {
          /* ignore */
        }
      }
    }

    var inject = ["slots"];
    exports.apply = apply;
    exports.inject = inject;
    exports.Panel = Panel;
    exports.Trigger = Trigger;
    return module.exports;
  },
});
