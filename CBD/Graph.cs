﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBD {
    public class NodeList<T> : Collection<Node<T>> {
        public NodeList() : base() { }
        public NodeList(int initialSize) {
            // Add the specified number of items
            for (int i = 0; i < initialSize; i++)
                base.Items.Add(default(Node<T>));
        }
        public Node<T> FindByValue(T value) {
            // search the list for the value
            foreach (Node<T> node in Items)
                if (node.Value.Equals(value))
                    return node;
            // if we reached here, we didn't find a matching node
            return null;
        }
    }
    public class Node<T> {
        // Private member-variables
        private T data;
        private NodeList<T> neighbors = null;
        public Node() { }
        public Node(T data) : this(data, null) { }
        public Node(T data, NodeList<T> neighbors) {
            this.data = data;
            this.neighbors = neighbors;
        }
        public T Value {
            get { return data; }
            set { data = value; }
        }
        protected NodeList<T> Neighbors {
            get { return neighbors; }
            set { neighbors = value; }
        }
        public bool HasNeighbor(T value) {
            foreach (Node<T> i in Neighbors) {
                if (i.Value.Equals(value)) return true;
            }
            return false;
        }
    }
    public class GraphNode<T> : CBD.Node<T> {
        private List<double> costs;
        public GraphNode() : base() { }
        public GraphNode(T value) : base(value) { }
        public GraphNode(T value, CBD.NodeList<T> neighbors) : base(value, neighbors) { }
        new public CBD.NodeList<T> Neighbors {
            get {
                if (base.Neighbors == null)
                    base.Neighbors = new CBD.NodeList<T>();
                return base.Neighbors;
            }
        }
        public List<double> Costs {
            get {
                if (costs == null)
                    costs = new List<double>();
                return costs;
            }
        }
    }
    public class Graph<T> : IEnumerable<T> {
        private NodeList<T> nodeSet;
        public Graph() : this(null) { }
        public Graph(NodeList<T> nodeSet) {
            if (nodeSet == null)
                this.nodeSet = new NodeList<T>();
            else
                this.nodeSet = nodeSet;
        }
        public NodeList<T> GetNodeSet() { return nodeSet; }
        public void AddNode(GraphNode<T> node)  {
            // adds a node to the graph
            nodeSet.Add(node);
        }
        public void AddNode(T value) {
            // adds a node to the graph
            nodeSet.Add(new GraphNode<T>(value));
        }
        public void AddDirectedEdge(GraphNode<T> from, GraphNode<T> to, double cost) {
            from.Neighbors.Add(to);
            from.Costs.Add(cost);
        }
        public void AddUndirectedEdge(GraphNode<T> from, GraphNode<T> to, double cost) {
            from.Neighbors.Add(to);
            from.Costs.Add(cost);
            to.Neighbors.Add(from);
            to.Costs.Add(cost);
        }
        public bool Contains(T value) { return nodeSet.FindByValue(value) != null; }
        public bool Remove(T value) {
            // first remove the node from the nodeset
            GraphNode<T> nodeToRemove = (GraphNode<T>)nodeSet.FindByValue(value);
            if (nodeToRemove == null)
                return false;
            // otherwise, the node was found
            nodeSet.Remove(nodeToRemove);
            // enumerate through each node in the nodeSet, removing edges to this node
            foreach (GraphNode<T> gnode in nodeSet) {
                int index = gnode.Neighbors.IndexOf(nodeToRemove);
                if (index != -1) {
                    // remove the reference to the node and associated cost
                    gnode.Neighbors.RemoveAt(index);
                    gnode.Costs.RemoveAt(index);
                }
            }
            return true;
        }
        public IEnumerator<T> GetEnumerator() { throw new NotImplementedException(); }
        IEnumerator IEnumerable.GetEnumerator() { throw new NotImplementedException(); }
        public NodeList<T> Nodes { get { return nodeSet; } }
        public int Count { get { return nodeSet.Count; } }
    }

}

