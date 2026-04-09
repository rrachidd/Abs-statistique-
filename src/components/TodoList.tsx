import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import { Button } from '../../components/ui/button';
import { Input } from '../../components/ui/input';
import { Card, CardHeader, CardTitle, CardContent } from '../../components/ui/card';
import { toast } from 'sonner';
import { Trash2, Plus, CheckCircle2, Circle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface Todo {
  id: string;
  task: string;
  is_complete: boolean;
  user_id: string;
}

export const TodoList: React.FC<{ userId: string }> = ({ userId }) => {
  const [todos, setTodos] = useState<Todo[]>([]);
  const [newTask, setNewTask] = useState('');
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchTodos();

    // Set up real-time subscription
    const subscription = supabase
      .channel('public:todos')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'todos' }, (payload) => {
        if (payload.eventType === 'INSERT') {
          setTodos((prev) => [payload.new as Todo, ...prev]);
        } else if (payload.eventType === 'DELETE') {
          setTodos((prev) => prev.filter((todo) => todo.id !== payload.old.id));
        } else if (payload.eventType === 'UPDATE') {
          setTodos((prev) =>
            prev.map((todo) => (todo.id === payload.new.id ? (payload.new as Todo) : todo))
          );
        }
      })
      .subscribe();

    return () => {
      supabase.removeChannel(subscription);
    };
  }, []);

  const fetchTodos = async () => {
    try {
      const { data, error } = await supabase
        .from('todos')
        .select('*')
        .order('created_at', { ascending: false });

      if (error) throw error;
      setTodos(data || []);
    } catch (error: any) {
      console.error('Error fetching todos:', error.message);
    } finally {
      setLoading(false);
    }
  };

  const addTodo = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTask.trim()) return;

    try {
      const { error } = await supabase
        .from('todos')
        .insert([{ task: newTask, user_id: userId, is_complete: false }]);

      if (error) throw error;
      setNewTask('');
      toast.success('Task added!');
    } catch (error: any) {
      toast.error(error.message);
    }
  };

  const toggleTodo = async (id: string, is_complete: boolean) => {
    try {
      const { error } = await supabase
        .from('todos')
        .update({ is_complete: !is_complete })
        .eq('id', id);

      if (error) throw error;
    } catch (error: any) {
      toast.error(error.message);
    }
  };

  const deleteTodo = async (id: string) => {
    try {
      const { error } = await supabase.from('todos').delete().eq('id', id);
      if (error) throw error;
      toast.success('Task deleted');
    } catch (error: any) {
      toast.error(error.message);
    }
  };

  return (
    <Card className="w-full max-w-2xl mx-auto">
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <CheckCircle2 className="text-primary" />
          My Tasks
        </CardTitle>
      </CardHeader>
      <CardContent className="space-y-6">
        <form onSubmit={addTodo} className="flex gap-2">
          <Input
            placeholder="What needs to be done?"
            value={newTask}
            onChange={(e) => setNewTask(e.target.value)}
          />
          <Button type="submit">
            <Plus className="w-4 h-4 mr-2" />
            Add
          </Button>
        </form>

        <div className="space-y-2">
          {loading ? (
            <p className="text-center text-muted-foreground py-8">Loading tasks...</p>
          ) : todos.length === 0 ? (
            <p className="text-center text-muted-foreground py-8">No tasks yet. Add one above!</p>
          ) : (
            <AnimatePresence mode="popLayout">
              {todos.map((todo) => (
                <motion.div
                  key={todo.id}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, x: -20 }}
                  className="flex items-center justify-between p-3 border rounded-lg bg-card hover:bg-accent/50 transition-colors"
                >
                  <div className="flex items-center gap-3">
                    <button
                      onClick={() => toggleTodo(todo.id, todo.is_complete)}
                      className="text-primary hover:scale-110 transition-transform"
                    >
                      {todo.is_complete ? (
                        <CheckCircle2 className="w-5 h-5" />
                      ) : (
                        <Circle className="w-5 h-5" />
                      )}
                    </button>
                    <span className={todo.is_complete ? 'line-through text-muted-foreground' : ''}>
                      {todo.task}
                    </span>
                  </div>
                  <Button
                    variant="ghost"
                    size="icon"
                    onClick={() => deleteTodo(todo.id)}
                    className="text-destructive hover:text-destructive hover:bg-destructive/10"
                  >
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </motion.div>
              ))}
            </AnimatePresence>
          )}
        </div>
      </CardContent>
    </Card>
  );
};
