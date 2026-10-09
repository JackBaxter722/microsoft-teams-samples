import React, { useCallback, useEffect, useMemo, useState } from "react";
import {
  Badge,
  Button,
  Card,
  CardHeader,
  Combobox,
  Divider,
  Field,
  FluentProvider,
  Input,
  Option,
  Spinner,
  Text,
  Textarea,
  teamsLightTheme
} from "@fluentui/react-components";

const stageLabels = {
  "needs-review": "Needs review",
  "candidate-unused": "Candidate unused",
  orphaned: "Orphaned"
};

const decisionLabels = {
  keep: "Keep",
  "delete-request": "Delete request",
  "reassign-request": "Reassign request",
  "quarantine-request": "Quarantine request"
};

export function App({ token, teams }) {
  const [session, setSession] = useState(null);
  const [resources, setResources] = useState([]);
  const [reviews, setReviews] = useState([]);
  const [skipToken, setSkipToken] = useState("");
  const [query, setQuery] = useState("");
  const [maker, setMaker] = useState(null);
  const [selectedResource, setSelectedResource] = useState("");
  const [stage, setStage] = useState("needs-review");
  const [rationale, setRationale] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const api = useCallback(
    async (path, options = {}) => {
      const response = await fetch(path, {
        ...options,
        headers: {
          Authorization: `******
          "Content-Type": "application/json",
          ...options.headers
        }
      });
      const body = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(body.error || `Request failed (${response.status}).`);
      return body;
    },
    [token]
  );

  const refresh = useCallback(async () => {
    setBusy(true);
    setError("");
    try {
      const [sessionResult, resourceResult, reviewResult] = await Promise.all([
        api("/api/session"),
        api("/api/resources"),
        api("/api/reviews")
      ]);
      setSession(sessionResult);
      setResources(resourceResult.resources);
      setSkipToken(resourceResult.skipToken || "");
      setReviews(reviewResult.reviews);
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setBusy(false);
    }
  }, [api]);

  useEffect(() => {
    refresh();
  }, [refresh]);

  const filteredResources = useMemo(() => {
    const value = query.trim().toLowerCase();
    if (!value) return resources;
    return resources.filter((resource) =>
      [resource.displayName, resource.type, resource.environmentId]
        .filter(Boolean)
        .some((field) => field.toLowerCase().includes(value))
    );
  }, [query, resources]);

  const resourceOptions = useMemo(
    () => resources.map((resource) => `${resource.displayName} (${resource.type})`),
    [resources]
  );

  async function loadMore() {
    if (!skipToken) return;
    setBusy(true);
    try {
      const page = await api(`/api/resources?skipToken=${encodeURIComponent(skipToken)}`);
      setResources((current) => [...current, ...page.resources]);
      setSkipToken(page.skipToken || "");
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setBusy(false);
    }
  }

  async function selectMaker() {
    try {
      const people = await teams.people.selectPeople({
        title: "Select the maker",
        singleSelect: true
      });
      const person = people?.[0];
      if (!person?.id) {
        setMaker(null);
        return;
      }
      setMaker({ id: person.id, name: person.displayName || person.email || person.id });
    } catch (pickerError) {
      setError(`Could not open the Teams people picker: ${pickerError.message}`);
    }
  }

  async function createReview(event) {
    event.preventDefault();
    if (!maker || !selectedResource || !rationale.trim()) return;
    setBusy(true);
    setError("");
    setNotice("");
    try {
      const resource = resources.find(
        (item) => `${item.displayName} (${item.type})` === selectedResource
      );
      if (!resource) throw new Error("Select a resource from the current inventory.");
      const result = await api("/api/reviews", {
        method: "POST",
        body: JSON.stringify({
          resourceId: resource.id,
          makerId: maker.id,
          makerName: maker.name,
          lifecycleStage: stage,
          rationale
        })
      });
      setNotice(result.notification);
      setMaker(null);
      setSelectedResource("");
      setRationale("");
      await refresh();
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setBusy(false);
    }
  }

  async function respond(review, decision) {
    let reassignee;
    if (decision === "reassign-request") {
      try {
        const people = await teams.people.selectPeople({
          title: "Select the proposed new owner",
          singleSelect: true
        });
        const person = people?.[0];
        if (!person?.id || person.id === session?.userId) return;
        reassignee = { reassigneeId: person.id, reassigneeName: person.displayName || person.email };
      } catch (pickerError) {
        setError(`Could not open the Teams people picker: ${pickerError.message}`);
        return;
      }
    }

    const target = reassignee ? ` to ${reassignee.reassigneeName}` : "";
    const confirmed = window.confirm(
      `Record "${decisionLabels[decision]}"${target} for ${review.displayName}? This records a review response only; it will not change the resource.`
    );
    if (!confirmed) return;

    setBusy(true);
    setError("");
    try {
      await api(`/api/reviews/${encodeURIComponent(review.id)}/decision`, {
        method: "POST",
        body: JSON.stringify({ decision, ...reassignee })
      });
      setNotice("Your response was recorded. No Power Platform resource was changed.");
      await refresh();
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <FluentProvider theme={teamsLightTheme}>
      <main className="page">
        <header className="header">
          <div>
            <Text as="h1" size={700} weight="semibold">Power Platform COE</Text>
            <Text as="p" size={300}>Inventory and lifecycle review</Text>
          </div>
          <Button appearance="secondary" onClick={refresh} disabled={busy}>Refresh</Button>
        </header>

        <div className="notice">
          Inventory is read-only here. “Candidate unused” and “Orphaned” are
          administrator-assessed review labels; this sample does not infer usage
          or execute delete, quarantine, or ownership changes.
        </div>

        {error && <div className="alert" role="alert">{error}</div>}
        {notice && <div className="notice success" role="status">{notice}</div>}
        {busy && <Spinner label="Loading Power Platform data…" />}

        <section className="section">
          <div className="section-heading">
            <Text as="h2" size={500} weight="semibold">Your visible resources</Text>
            {session?.isAdministrator && <Badge appearance="filled" color="brand">Administrator</Badge>}
          </div>
          <Field label="Search resources">
            <Combobox
              freeform
              placeholder="Search by name, type, or environment"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
            >
              {filteredResources.slice(0, 30).map((resource) => (
                <Option key={resource.id} value={resource.displayName}>
                  {resource.displayName} · {resource.type}
                </Option>
              ))}
            </Combobox>
          </Field>
          <p className="muted">
            {resources.length} resource(s) on this page. Results are limited to what
            your delegated Inventory API access can see.
          </p>
          <div className="resource-grid">
            {filteredResources.map((resource) => {
              const relatedReviews = reviews.filter((review) => review.resourceId === resource.id);
              return (
                <Card key={resource.id} className="resource-card">
                  <CardHeader
                    header={<Text weight="semibold">{resource.displayName}</Text>}
                    description={<Text size={200}>{resource.type}</Text>}
                  />
                  <div className="card-content">
                    <Text size={200}>Resource ID: {resource.id}</Text>
                    <Text size={200}>Environment: {resource.environmentId || "Not reported"}</Text>
                    <Text size={200}>Owner ID: {resource.ownerId || "Not reported"}</Text>
                    <Text size={200}>Created: {resource.createdAt || "Not reported"}</Text>
                    <Text size={200}>Last modified: {resource.lastModifiedAt || "Not reported"}</Text>
                    {relatedReviews.map((review) => (
                      <div className="review-chip" key={review.id}>
                        <Badge color={review.status === "pending" ? "warning" : "success"}>
                          {stageLabels[review.lifecycleStage]} · {review.status}
                        </Badge>
                        {review.decisions.map((item, index) => (
                          <Text size={200} key={`${item.at}-${index}`}>
                            {decisionLabels[item.decision]} · {new Date(item.at).toLocaleString()}
                            {item.reassigneeName ? ` · proposed owner: ${item.reassigneeName}` : ""}
                          </Text>
                        ))}
                      </div>
                    ))}
                  </div>
                </Card>
              );
            })}
          </div>
          {filteredResources.length === 0 && !busy && <p>No matching resources found.</p>}
          {skipToken && <Button appearance="secondary" onClick={loadMore} disabled={busy}>Load more</Button>}
        </section>

        <Divider />

        <section className="section">
          <Text as="h2" size={500} weight="semibold">Review decisions</Text>
          <div className="review-grid">
            {reviews.map((review) => (
              <Card key={review.id} className="review-card">
                <CardHeader
                  header={<Text weight="semibold">{review.displayName}</Text>}
                  description={<Text size={200}>{stageLabels[review.lifecycleStage]} · {review.status}</Text>}
                />
                <div className="card-content">
                  <Text size={200}>Stage rationale: {review.stageRationale}</Text>
                  <Text size={200}>Sent: {new Date(review.createdAt).toLocaleString()}</Text>
                  {review.decisions.map((decision, index) => (
                    <Text size={200} key={`${decision.at}-${index}`}>
                      {decisionLabels[decision.decision]} · {new Date(decision.at).toLocaleString()}
                      {decision.reassigneeName ? ` · proposed owner: ${decision.reassigneeName}` : ""}
                    </Text>
                  ))}
                  {review.status === "pending" && !session?.isAdministrator && (
                    <div className="actions">
                      {["keep", "delete-request", "reassign-request", "quarantine-request"].map((decision) => (
                        <Button
                          key={decision}
                          appearance={decision === "keep" ? "primary" : "secondary"}
                          disabled={busy}
                          onClick={() => respond(review, decision)}
                        >
                          {decisionLabels[decision]}
                        </Button>
                      ))}
                    </div>
                  )}
                </div>
              </Card>
            ))}
          </div>
          {reviews.length === 0 && <p>No reviews available.</p>}
        </section>

        {session?.isAdministrator && (
          <>
            <Divider />
            <section className="section">
              <Text as="h2" size={500} weight="semibold">Send a maker review</Text>
              <form className="admin-form" onSubmit={createReview}>
                <Field label="Inventory resource" required>
                  <Combobox
                    placeholder="Choose a resource"
                    value={selectedResource}
                    onOptionSelect={(_, data) => setSelectedResource(data.optionValue || "")}
                    onChange={(event) => setSelectedResource(event.target.value)}
                  >
                    {resourceOptions.map((option) => <Option key={option} value={option}>{option}</Option>)}
                  </Combobox>
                </Field>
                <Field label="Maker" required>
                  <div className="picker">
                    <Input readOnly value={maker?.name || ""} placeholder="Select a maker in Teams" />
                    <Button appearance="secondary" type="button" onClick={selectMaker}>Choose person</Button>
                  </div>
                </Field>
                <Field label="Administrator-assessed stage" required>
                  <select value={stage} onChange={(event) => setStage(event.target.value)}>
                    <option value="needs-review">Needs review</option>
                    <option value="candidate-unused">Candidate unused</option>
                    <option value="orphaned">Orphaned</option>
                  </select>
                </Field>
                <Field label="Review rationale" required>
                  <Textarea
                    value={rationale}
                    maxLength={500}
                    onChange={(event) => setRationale(event.target.value)}
                    placeholder="Explain the assessment and what the maker should review."
                  />
                </Field>
                <Button
                  appearance="primary"
                  type="submit"
                  disabled={busy || !maker || !selectedResource || !rationale.trim()}
                >
                  Send review card
                </Button>
              </form>
            </section>
          </>
        )}
      </main>
    </FluentProvider>
  );
}
