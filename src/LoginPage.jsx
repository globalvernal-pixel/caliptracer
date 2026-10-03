import React, { useState, useEffect } from 'react';
import { Lock, User, Eye, EyeOff, Loader2 } from 'lucide-react';
import logo from './logo.png';

export default function LoginPage({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [shake, setShake] = useState(false);

  const triggerShake = () => {
    setShake(true);
    setTimeout(() => setShake(false), 500);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password.trim()) {
      setError('Please enter username and password.');
      triggerShake();
      return;
    }
    setLoading(true);
    setError('');
    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: username.trim(), password }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || 'Login failed. Please try again.');
        triggerShake();
      } else {
        localStorage.setItem('caliph_token', data.token);
        localStorage.setItem('caliph_user', JSON.stringify(data.user));
        onLogin(data.user);
      }
    } catch {
      setError('Connection error. Please check your network.');
      triggerShake();
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen w-full flex items-center justify-center relative overflow-hidden" style={{background: 'linear-gradient(135deg, #0f172a 0%, #1e3a5f 40%, #0f2340 100%)'}}>
      {/* Animated background orbs */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-40 -left-40 w-96 h-96 rounded-full opacity-20" style={{background: 'radial-gradient(circle, #3b82f6, transparent)', animation: 'pulse 4s ease-in-out infinite'}} />
        <div className="absolute -bottom-32 -right-32 w-80 h-80 rounded-full opacity-15" style={{background: 'radial-gradient(circle, #6366f1, transparent)', animation: 'pulse 6s ease-in-out infinite 2s'}} />
        <div className="absolute top-1/2 left-1/4 w-64 h-64 rounded-full opacity-10" style={{background: 'radial-gradient(circle, #0ea5e9, transparent)', animation: 'pulse 5s ease-in-out infinite 1s'}} />
      </div>

      {/* Grid overlay */}
      <div className="absolute inset-0 opacity-5" style={{backgroundImage: 'linear-gradient(rgba(255,255,255,0.1) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.1) 1px, transparent 1px)', backgroundSize: '40px 40px'}} />

      {/* Login card */}
      <div
        className={`relative w-full max-w-sm mx-4 ${shake ? 'animate-shake' : ''}`}
        style={{
          animation: shake ? 'shake 0.4s ease' : 'slideUp 0.5s ease-out',
        }}
      >
        {/* Glass card */}
        <div className="relative rounded-3xl overflow-hidden" style={{background: 'rgba(255,255,255,0.07)', backdropFilter: 'blur(24px)', border: '1px solid rgba(255,255,255,0.15)', boxShadow: '0 32px 64px rgba(0,0,0,0.4)'}}>
          {/* Top accent line */}
          <div className="h-1 w-full" style={{background: 'linear-gradient(90deg, #3b82f6, #6366f1, #0ea5e9)'}} />

          <div className="p-8 flex flex-col items-center gap-6">
            {/* Logo + Title */}
            <div className="flex flex-col items-center gap-3">
              <div className="w-20 h-20 rounded-2xl overflow-hidden shadow-2xl flex items-center justify-center" style={{background: 'rgba(255,255,255,0.1)', border: '1px solid rgba(255,255,255,0.2)'}}>
                <img src={logo} alt="Caliph Logo" className="w-16 h-16 object-contain" onError={(e) => { e.target.style.display='none'; }} />
              </div>
              <div className="text-center">
                <h1 className="text-2xl font-black text-white tracking-tight">CALIPH</h1>
                <p className="text-xs font-medium mt-1" style={{color: 'rgba(148,163,184,0.9)'}}>Student Management System</p>
              </div>
            </div>

            {/* Divider */}
            <div className="w-full flex items-center gap-3">
              <div className="flex-1 h-px" style={{background: 'rgba(255,255,255,0.1)'}} />
              <span className="text-xs font-semibold" style={{color: 'rgba(148,163,184,0.6)'}}>SIGN IN TO CONTINUE</span>
              <div className="flex-1 h-px" style={{background: 'rgba(255,255,255,0.1)'}} />
            </div>

            {/* Form */}
            <form onSubmit={handleSubmit} className="w-full flex flex-col gap-4">
              {/* Username */}
              <div className="relative">
                <div className="absolute left-4 top-1/2 -translate-y-1/2" style={{color: 'rgba(148,163,184,0.7)'}}>
                  <User className="w-4 h-4" />
                </div>
                <input
                  id="login-username"
                  type="text"
                  placeholder="Username"
                  value={username}
                  onChange={(e) => { setUsername(e.target.value); setError(''); }}
                  autoComplete="username"
                  autoFocus
                  className="w-full pl-11 pr-4 py-3.5 rounded-xl text-sm font-semibold text-white placeholder-slate-400 outline-none transition-all"
                  style={{
                    background: 'rgba(255,255,255,0.08)',
                    border: '1px solid rgba(255,255,255,0.12)',
                  }}
                  onFocus={(e) => { e.target.style.border = '1px solid rgba(99,102,241,0.7)'; e.target.style.background = 'rgba(255,255,255,0.12)'; }}
                  onBlur={(e) => { e.target.style.border = '1px solid rgba(255,255,255,0.12)'; e.target.style.background = 'rgba(255,255,255,0.08)'; }}
                />
              </div>

              {/* Password */}
              <div className="relative">
                <div className="absolute left-4 top-1/2 -translate-y-1/2" style={{color: 'rgba(148,163,184,0.7)'}}>
                  <Lock className="w-4 h-4" />
                </div>
                <input
                  id="login-password"
                  type={showPassword ? 'text' : 'password'}
                  placeholder="Password"
                  value={password}
                  onChange={(e) => { setPassword(e.target.value); setError(''); }}
                  autoComplete="current-password"
                  className="w-full pl-11 pr-12 py-3.5 rounded-xl text-sm font-semibold text-white placeholder-slate-400 outline-none transition-all"
                  style={{
                    background: 'rgba(255,255,255,0.08)',
                    border: '1px solid rgba(255,255,255,0.12)',
                  }}
                  onFocus={(e) => { e.target.style.border = '1px solid rgba(99,102,241,0.7)'; e.target.style.background = 'rgba(255,255,255,0.12)'; }}
                  onBlur={(e) => { e.target.style.border = '1px solid rgba(255,255,255,0.12)'; e.target.style.background = 'rgba(255,255,255,0.08)'; }}
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-4 top-1/2 -translate-y-1/2 transition-opacity hover:opacity-100"
                  style={{color: 'rgba(148,163,184,0.7)'}}
                >
                  {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>

              {/* Error */}
              {error && (
                <div className="flex items-center gap-2 px-3 py-2.5 rounded-xl text-xs font-bold" style={{background: 'rgba(239,68,68,0.15)', border: '1px solid rgba(239,68,68,0.3)', color: '#fca5a5'}}>
                  <span className="w-1.5 h-1.5 rounded-full bg-red-400 shrink-0" />
                  {error}
                </div>
              )}

              {/* Submit */}
              <button
                id="login-submit"
                type="submit"
                disabled={loading}
                className="w-full py-3.5 rounded-xl font-extrabold text-sm text-white tracking-wider uppercase transition-all active:scale-95 flex items-center justify-center gap-2 mt-1"
                style={{
                  background: loading ? 'rgba(99,102,241,0.5)' : 'linear-gradient(135deg, #4f46e5, #3b82f6)',
                  boxShadow: loading ? 'none' : '0 8px 24px rgba(79,70,229,0.4)',
                  cursor: loading ? 'not-allowed' : 'pointer',
                }}
              >
                {loading ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    Signing In...
                  </>
                ) : (
                  <>
                    <Lock className="w-4 h-4" />
                    Sign In
                  </>
                )}
              </button>
            </form>
          </div>

          {/* Footer */}
          <div className="px-8 pb-6 text-center">
            <p className="text-xs" style={{color: 'rgba(100,116,139,0.8)'}}>
              Caliph School Management · Secure Access
            </p>
          </div>
        </div>
      </div>

      <style>{`
        @keyframes slideUp {
          from { opacity: 0; transform: translateY(24px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes shake {
          0%, 100% { transform: translateX(0); }
          20% { transform: translateX(-8px); }
          40% { transform: translateX(8px); }
          60% { transform: translateX(-6px); }
          80% { transform: translateX(6px); }
        }
        @keyframes pulse {
          0%, 100% { opacity: 0.15; transform: scale(1); }
          50% { opacity: 0.25; transform: scale(1.05); }
        }
        input::placeholder { color: rgba(148,163,184,0.5); }
      `}</style>
    </div>
  );
}
