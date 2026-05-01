import React, { useEffect, useState } from "react";
import { MessageSquarePlus, RefreshCw, Send } from "lucide-react";
import { API } from "../constants";

export function FeedbackPage() {
  const [items, setItems] = useState([]);
  const [message, setMessage] = useState("");
  const [author, setAuthor] = useState("");
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");

  async function loadFeedback() {
    setLoading(true);
    setError("");
    setSuccess("");

    try {
      const response = await fetch(`${API}/api/feedback`);
      const data = await response.json();

      if (!response.ok || !data.success) {
        setError(data.message || "Unable to load feedback.");
        return;
      }

      setItems(data.items || []);
    } catch (err) {
      setError(err.message || "Unable to load feedback.");
    } finally {
      setLoading(false);
    }
  }

  async function submitFeedback(event) {
    event.preventDefault();
    setSaving(true);
    setError("");
    setSuccess("");

    try {
      const response = await fetch(`${API}/api/feedback`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message, author })
      });
      const data = await response.json();

      if (!response.ok || !data.success) {
        setError(data.message || "Unable to save feedback.");
        return;
      }

      setItems((current) => [data.item, ...current]);
      setMessage("");
      setAuthor("");
      setSuccess("Feedback saved.");
    } catch (err) {
      setError(err.message || "Unable to save feedback.");
    } finally {
      setSaving(false);
    }
  }

  useEffect(() => {
    loadFeedback();
  }, []);

  return (
    <section className="page">
      <header className="page-head">
        <div>
          <h1>Feedback</h1>
          <p>Collect feature ideas and change requests from WorkSight users.</p>
        </div>
        <button className="ghost-btn" onClick={loadFeedback} disabled={loading}>
          <RefreshCw size={16} /> Refresh
        </button>
      </header>

      <div className="feedback-layout">
        <form className="panel feedback-form" onSubmit={submitFeedback}>
          <div className="panel-title">
            <MessageSquarePlus size={20} />
            <span>New Feedback</span>
          </div>

          <label className="ws-field">
            <span>Content</span>
            <textarea
              value={message}
              onChange={(event) => setMessage(event.target.value)}
              placeholder="Write the feature you want or the workflow that needs improvement"
            />
          </label>

          <label className="ws-field">
            <span>Name</span>
            <input value={author} onChange={(event) => setAuthor(event.target.value)} placeholder="Optional" />
          </label>

          <button className="primary-btn feedback-submit" type="submit" disabled={saving}>
            <Send size={16} /> {saving ? "Saving" : "Submit Feedback"}
          </button>

          {error && <div className="error">{error}</div>}
          {success && <div className="success">{success}</div>}
        </form>

        <div className="panel feedback-list-panel">
          <div className="table-head">
            <h2>Recent Feedback</h2>
            <span className="hint-line">{loading ? "Loading..." : `${items.length} items`}</span>
          </div>

          {!items.length && !loading && <div className="empty">No feedback yet.</div>}

          <div className="feedback-list">
            {items.map((item) => (
              <article className="feedback-item" key={item.id}>
                <div className="feedback-item-head">
                  <span className="feedback-badge feature">Feedback</span>
                  <time>{new Date(item.createdAt).toLocaleString()}</time>
                </div>
                {item.title && <h2>{item.title}</h2>}
                <p>{item.message}</p>
                {item.author && <small>By {item.author}</small>}
              </article>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}
