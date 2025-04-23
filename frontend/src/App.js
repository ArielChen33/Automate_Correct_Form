import logo from './logo.svg';
import React, { useEffect, useState } from "react";

import './App.css';

function App() {

  const [message, setMessage] = useState("");

  useEffect(() => {
    fetch("http://localhost:8000")
    .then(res => res.json())
    .then(data => setMessage(data)
    )
  }, [])

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          This is just for testing purpose!!
        </p>
        <h1>hiii {message}</h1>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
      </header>
    </div>
  );
}

export default App;
