import React, { useEffect, useState } from 'react';
import { auth, onAuthStateChanged, signOut } from './firebase';
import { Button } from './components/ui/button';
import { Toaster } from 'sonner';
import { LogOut, Database, Shield, HardDrive, Github } from 'lucide-react';
import { motion } from 'motion/react';

function App() {
  const [user, setUser] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (user) => {
      setUser(user);
      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  const handleSignOut = async () => {
    await signOut(auth);
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
            <h1 className="text-xl font-bold tracking-tight">Firebase Absence Manager</h1>
          </div>
          
          {user && (
            <div className="flex items-center gap-4">
              <span className="text-sm text-muted-foreground hidden sm:inline-block">
                {user.email}
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
        {!user ? (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="max-w-4xl mx-auto space-y-12"
          >
            <div className="text-center space-y-4">
              <h2 className="text-4xl font-extrabold tracking-tight sm:text-6xl">
                School <span className="text-primary">Absence Manager</span>
              </h2>
              <p className="text-xl text-muted-foreground max-w-2xl mx-auto">
                Manage student absences, generate reports, and communicate with parents via WhatsApp.
              </p>
            </div>

            <div className="flex justify-center">
               <p className="text-muted-foreground">Please use the login button in the header or sidebar to continue.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 pt-12">
              <FeatureCard 
                icon={<Shield className="w-6 h-6" />}
                title="Authentication"
                description="Secure login with Google to keep your data safe."
              />
              <FeatureCard 
                icon={<Database className="w-6 h-6" />}
                title="Cloud Storage"
                description="Sync your datasets across devices automatically."
              />
              <FeatureCard 
                icon={<HardDrive className="w-6 h-6" />}
                title="Local Persistence"
                description="Work offline and sync when you're back online."
              />
            </div>
          </motion.div>
        ) : (
          <div className="max-w-4xl mx-auto space-y-8">
             <div className="bg-card border rounded-2xl p-8 text-center space-y-4">
                <h3 className="text-2xl font-bold">Welcome back!</h3>
                <p className="text-muted-foreground">
                  Your data is being synced with Firebase. You can manage your classes and students in the main interface.
                </p>
             </div>
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
