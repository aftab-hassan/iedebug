import React from 'react';

interface ErrorBoundaryState {
    hasError: boolean;
  }

export class ErrorBoundary extends React.Component<{}, ErrorBoundaryState> {
    constructor(props: {}) {
      super(props);
      this.state = { hasError: false };
    }

    static getDerivedStateFromError(error: any) {
      // Update state so the next render will show the fallback UI.
      return { hasError: true };
    }

    componentDidCatch(error: any, info: any) {
        alert('printing error: ' + error);
        alert('printing info: ' + info);

      // You can also log the error to an error reporting service
      console.log('printing error: ' + error);
      console.log('printing info: ' + info);
    }

    render() {
      if (this.state.hasError) {
        // You can render any custom fallback UI
        return <h1>Something went wrong.</h1>;
      }

      return this.props.children;
    }
  }