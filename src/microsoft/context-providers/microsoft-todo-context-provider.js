import React, { useCallback, useEffect, useMemo, useState } from 'react';
import PropTypes from 'prop-types';
import log from 'loglevel';
import format from 'string-template';
import { useAuth } from '../../context-hooks/auth-context-hooks';
import { ToDoContext } from '../../context-hooks/todo-context-hooks';

const logger = log.getLogger('Microsoft');
const refreshInterval = 60000;

// Override with env var in your webpack define
// Template placeholders: {id}, {listId}. Falls back to To Do web
const todoTaskTemplateUrl =
  process.env.TODO_TASK_TEMPLATE_URL || 'https://to-do.office.com/tasks/';

function sortTasks(a, b) {
  // Sort by due date (nulls last), then importance (high > normal > low)
  const impRank = (v) => ({ high: 0, normal: 1, low: 2 }[v] ?? 1);

  const aDue = a.due ? new Date(a.due).getTime() : Number.POSITIVE_INFINITY;
  const bDue = b.due ? new Date(b.due).getTime() : Number.POSITIVE_INFINITY;

  if (aDue !== bDue) {
    return aDue - bDue;
  }
  return impRank(a.importance) - impRank(b.importance);
}

export function MicrosoftToDoProvider({ children }) {
  const { client, loggedIn } = useAuth();

  const [error, setError] = useState(false);
  const [state, setState] = useState('load');
  const [tasks, setTasks] = useState([]);
  const [renderCount, setRenderCount] = useState(0);

  const mapTask = useCallback((t, list) => {
    // Construct task link
    let link;
    if (todoTaskTemplateUrl.includes('{id}')) {
      link = format(todoTaskTemplateUrl, { id: t.id, listId: list.id });
    } else {
      link = todoTaskTemplateUrl;
    }

    return {
      id: t.id,
      title: t.title,
      status: t.status,
      importance: t.importance,
      due: t.dueDateTime?.dateTime || null,
      created: t.createdDateTime,
      modified: t.lastModifiedDateTime,
      completed: t.completedDateTime?.dateTime || null,
      listId: list.id,
      listName: list.displayName,
      link
    };
  }, []);

  const fetchTasks = useCallback(async () => {
        if (!loggedIn) {
            return;
        }
        if (state === 'refresh' && document.hidden) {
            return;
        }

        logger.debug(`${tasks.length === 0 ? 'loading' : 'refreshing'} Microsoft To Do tasks`);

        try {
            // Get all lists (safe select + fallback)
            let listsResp;
            try {
                listsResp = await client
                    .api('/me/todo/lists')
                    // avoid wellknownListName to prevent 400s
                    //  .select(['id', 'displayName'])
                    .get();
            } catch (err) {
                logger.warn('Lists with $select failed, retrying without $select', err?.statusCode || '', err?.message || '');
                listsResp = await client.api('/me/todo/lists').get();
            }

            const lists = listsResp?.value ?? [];
            if (lists.length === 0) {
                setTasks([]);
                setError(false);
                setState('loaded');
                setRenderCount(c => c + 1);
                return;
            }

            const perList = await Promise.all(lists.map(async (list) => {
                try {
                    const tr = await client
                        .api(`/me/todo/lists/${list.id}/tasks`)
                        .get();

                    const items = (tr?.value ?? []).map(t => ({
                        id: t.id,
                        title: t.title,
                        status: t.status,
                        importance: t.importance,
                        due: t.dueDateTime?.dateTime || null,
                        created: t.createdDateTime,
                        modified: t.lastModifiedDateTime,
                        completed: t.completedDateTime?.dateTime || null,
                        listId: list.id,
                        listName: list.displayName,
                        link: todoTaskTemplateUrl.includes('{id}')
                            ? format(todoTaskTemplateUrl, { id: t.id, listId: list.id })
                            : todoTaskTemplateUrl
                    }));
                    return items;
                } catch (e) {
                    logger.error('Error fetching tasks for list', list.id, e?.statusCode || '', e?.message || '');
                    return [];
                }
            }));

            const merged = perList.flat().sort(sortTasks).slice(0, 20);

            setTasks(merged);
            setError(false);
            setState('loaded');
            setRenderCount(c => c + 1);
        } catch (e) {
            const detail = e?.body || e?.message || JSON.stringify(e);
            logger.error('Error loading Microsoft To Do tasks', detail);
            setError(true);
            setState('loaded');
            setRenderCount(c => c + 1);
        }
    }, [client, loggedIn, state, tasks.length, mapTask]);

  const toggleComplete = useCallback(
    async (task) => {
      const currentlyCompleted = task.status === 'completed';
      const nextStatus = currentlyCompleted ? 'notStarted' : 'completed';

      // Optimistic update
      setTasks((prev) =>
        prev.map((t) => (t.id === task.id ? { ...t, status: nextStatus } : t))
      );

      try {
        await client
          .api(`/me/todo/lists/${task.listId}/tasks/${task.id}`)
          .header('Content-Type', 'application/json')
          .patch({ status: nextStatus });

        // Refresh to reconcile server truth
        setState('refresh');
      } catch (e) {
        logger.error('Error updating Microsoft To Do task', e);
        // Revert on failure
        setTasks((prev) =>
          prev.map((t) => (t.id === task.id ? { ...t, status: task.status } : t))
        );
        setError(true);
      }
    },
    [client]
  );

  useEffect(() => {
    if (loggedIn && (state === 'load' || state === 'refresh')) {
      fetchTasks();
    }

    if (!loggedIn && state === 'loaded') {
      setTasks([]);
      setState('load');
    }
  }, [loggedIn, state, fetchTasks]);

  useEffect(() => {
    let timerId;

    const startInterval = () => {
      stopInterval();
      if (!document.hidden) {
        timerId = setInterval(() => setState('refresh'), refreshInterval);
      }
    };

    const stopInterval = () => {
      if (timerId) {
        clearInterval(timerId);
        timerId = undefined;
      }
    };

    const onVisibilityChange = () => {
      logger.debug('Microsoft To Do visibility changed');
      if (document.hidden) {
        stopInterval();
      } else {
        setState((s) => {
          if (s === 'loaded') {
            return 'refresh';
          }
          return s;
        });
        startInterval();
      }
    };

    if (loggedIn) {
      document.addEventListener('visibilitychange', onVisibilityChange);
      startInterval();
    }

    return () => {
      document.removeEventListener('visibilitychange', onVisibilityChange);
      stopInterval();
    };
  }, [loggedIn]);

  const contextValue = useMemo(
    () => ({
      error,
      tasks,
      refresh: () => setState('refresh'),
      toggleComplete,
      state
    }),
    [error, tasks, state, toggleComplete, renderCount]
  );

  useEffect(() => {
    logger.debug('MicrosoftToDoProvider mounted');
    return () => logger.debug('MicrosoftToDoProvider unmounted');
  }, []);

  return <ToDoContext.Provider value={contextValue}>{children}</ToDoContext.Provider>;
}

MicrosoftToDoProvider.propTypes = {
  children: PropTypes.node.isRequired
};
