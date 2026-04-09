import React, { useEffect, useState } from 'react';
import { supabase } from './lib/supabase';
import { Auth } from './components/Auth';
import { TodoList } from './components/TodoList';
import { FileStorage } from './components/FileStorage';
import { Button } from '../components/ui/button';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '../components/ui/tabs';
import { Toaster } from '../components/ui/sonner';
import { LogOut, Database, Shield, HardDrive, Github } from 'lucide-react';
import { motion } from 'motion/react';

function App() {
  const [session, setSession] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // Check current session
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setLoading(false);
    });

    // Listen for auth changes
    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });

    return () => subscription.unsubscribe();
  }, []);

  const handleSignOut = async () => {
    await supabase.auth.signOut();
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-background">
        <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-primary"></div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background text-foreground font-sans antialiased">
      <Toaster position="top-center" />
      
      <header className="border-b bg-card/50 backdrop-blur-md sticky top-0 z-50">
        <div className="container mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-primary p-1.5 rounded-lg">
              <Database className="w-5 h-5 text-primary-foreground" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">Supabase Starter</h1>
          </div>
          
          {session && (
            <div className="flex items-center gap-4">
              <span className="text-sm text-muted-foreground hidden sm:inline-block">
                {session.user.email}
              </span>
              <Button variant="outline" size="sm" onClick={handleSignOut}>
                <LogOut className="w-4 h-4 mr-2" />
                Sign Out
              </Button>
            </div>
          )}
        </div>
      </header>

      <main className="container mx-auto px-4 py-8">
        {!session ? (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="max-w-4xl mx-auto space-y-12"
          >
            <div className="text-center space-y-4">
              <h2 className="text-4xl font-extrabold tracking-tight sm:text-6xl">
                The Open Source <span className="text-primary">Firebase Alternative</span>
              </h2>
              <p className="text-xl text-muted-foreground max-w-2xl mx-auto">
                Build production-grade applications with a Postgres database, Authentication, 
                Instant APIs, Edge Functions, Realtime subscriptions, and Storage.
              </p>
            </div>

            <Auth />

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 pt-12">
              <FeatureCard 
                icon={<Shield className="w-6 h-6" />}
                title="Authentication"
                description="User management with social logins, magic links, and more."
              />
              <FeatureCard 
                icon={<Database className="w-6 h-6" />}
                title="Postgres Database"
                description="Every project comes with a full Postgres database."
              />
              <FeatureCard 
                icon={<HardDrive className="w-6 h-6" />}
                title="Storage"
                description="Store, organize, and serve large files like images and videos."
              />
            </div>
          </motion.div>
        ) : (
          <div className="max-w-4xl mx-auto space-y-8">
            <Tabs defaultValue="todos" className="w-full">
              <TabsList className="grid w-full grid-cols-2 mb-8">
                <TabsTrigger value="todos" className="flex items-center gap-2">
                  <Database className="w-4 h-4" />
                  Database & Realtime
                </TabsTrigger>
                <TabsTrigger value="storage" className="flex items-center gap-2">
                  <HardDrive className="w-4 h-4" />
                  Storage
                </TabsTrigger>
              </TabsList>
              
              <TabsContent value="todos">
                <motion.div
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  transition={{ duration: 0.3 }}
                >
                  <TodoList userId={session.user.id} />
                </motion.div>
              </TabsContent>
              
              <TabsContent value="storage">
                <motion.div
                  initial={{ opacity: 0, x: 20 }}
                  animate={{ opacity: 1, x: 0 }}
                  transition={{ duration: 0.3 }}
                >
                  <FileStorage userId={session.user.id} />
                </motion.div>
              </TabsContent>
            </Tabs>
          </div>
        )}
      </main>

      <footer className="border-t py-8 mt-auto">
        <div className="container mx-auto px-4 flex flex-col sm:flex-row justify-between items-center gap-4">
          <p className="text-sm text-muted-foreground">
            Built with React, Tailwind, and Supabase.
          </p>
          <div className="flex items-center gap-4">
            <a href="https://supabase.com" target="_blank" rel="noreferrer" className="text-muted-foreground hover:text-primary transition-colors">
              Documentation
            </a>
            <a href="https://github.com/supabase" target="_blank" rel="noreferrer" className="text-muted-foreground hover:text-primary transition-colors">
              <Github className="w-5 h-5" />
            </a>
          </div>
        </div>
      </footer>
    </div>
  );
}

function FeatureCard({ icon, title, description }: { icon: React.ReactNode, title: string, description: string }) {
  return (
    <div className="p-6 rounded-2xl border bg-card hover:shadow-lg transition-all duration-300 hover:-translate-y-1">
      <div className="w-12 h-12 rounded-xl bg-primary/10 flex items-center justify-center text-primary mb-4">
        {icon}
      </div>
      <h3 className="text-lg font-bold mb-2">{title}</h3>
      <p className="text-muted-foreground text-sm leading-relaxed">
        {description}
      </p>
    </div>
  );
}

export default App;
